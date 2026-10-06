from __future__ import annotations

import json
import os
import re
import sqlite3
from contextlib import contextmanager
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Iterable, Iterator, Sequence

from .parser import QuestionDraft
from .taxonomy import enrich_question_draft, infer_question_metadata


SQLITE_SCHEMA = """
CREATE TABLE IF NOT EXISTS imports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_filename TEXT NOT NULL,
    stored_filename TEXT NOT NULL,
    preview_relpath TEXT,
    stats_json TEXT NOT NULL DEFAULT '{}',
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS questions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_id INTEGER NOT NULL,
    number TEXT, question_type TEXT, category_major TEXT, category_minor TEXT, difficulty TEXT,
    curriculum_version TEXT NOT NULL DEFAULT '人教B版',
    curriculum_chapter TEXT, curriculum_section TEXT,
    classification_source TEXT NOT NULL DEFAULT 'rule',
    classification_confidence REAL, review_reason TEXT,
    asset_manifest TEXT NOT NULL DEFAULT '[]',
    stem TEXT NOT NULL, stem_html TEXT, options TEXT, options_html TEXT,
    answer TEXT, answer_html TEXT, analysis TEXT, analysis_html TEXT,
    source_name TEXT, raw_text TEXT, raw_html TEXT,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(import_id) REFERENCES imports(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_questions_import_id ON questions(import_id);
CREATE INDEX IF NOT EXISTS idx_questions_type ON questions(question_type);
CREATE INDEX IF NOT EXISTS idx_questions_category ON questions(category_major);
CREATE INDEX IF NOT EXISTS idx_questions_difficulty ON questions(difficulty);
CREATE TABLE IF NOT EXISTS papers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    title TEXT NOT NULL, exam_time TEXT NOT NULL DEFAULT '', total_score TEXT NOT NULL DEFAULT '',
    school TEXT NOT NULL DEFAULT '', grade TEXT NOT NULL DEFAULT '', class_name TEXT NOT NULL DEFAULT '',
    instructions TEXT NOT NULL DEFAULT '', source TEXT NOT NULL DEFAULT 'manual',
    status TEXT NOT NULL DEFAULT 'saved', created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS paper_questions (
    id INTEGER PRIMARY KEY AUTOINCREMENT, paper_id INTEGER NOT NULL,
    question_id INTEGER, original_question_id INTEGER NOT NULL,
    position INTEGER NOT NULL, display_number INTEGER NOT NULL,
    FOREIGN KEY(paper_id) REFERENCES papers(id) ON DELETE CASCADE,
    FOREIGN KEY(question_id) REFERENCES questions(id) ON DELETE SET NULL,
    UNIQUE(paper_id, original_question_id)
);
CREATE INDEX IF NOT EXISTS idx_paper_questions_paper ON paper_questions(paper_id, position);
CREATE TABLE IF NOT EXISTS export_records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    paper_id INTEGER NOT NULL,
    export_format TEXT NOT NULL,
    filename TEXT NOT NULL,
    question_count INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(paper_id) REFERENCES papers(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_export_records_created ON export_records(created_at DESC, id DESC);
"""

QUESTION_COLUMNS = {
    "category_major": "TEXT", "category_minor": "TEXT", "difficulty": "TEXT",
    "curriculum_version": "TEXT NOT NULL DEFAULT '人教B版'",
    "curriculum_chapter": "TEXT", "curriculum_section": "TEXT",
    "classification_source": "TEXT NOT NULL DEFAULT 'legacy'",
    "classification_confidence": "REAL", "review_reason": "TEXT",
    "asset_manifest": "TEXT NOT NULL DEFAULT '[]'",
    "stem_html": "TEXT", "options_html": "TEXT", "answer_html": "TEXT",
    "analysis_html": "TEXT", "raw_html": "TEXT",
}


@dataclass(frozen=True)
class DatabaseConfig:
    driver: str = "sqlite"
    sqlite_path: Path | None = None


def load_database_config(project_root: str | Path) -> DatabaseConfig:
    root = Path(project_root).resolve()
    env = dict(os.environ)
    env.update(_read_dotenv(root / ".env"))
    path = Path(env.get("SQLITE_PATH", "data/question_bank.db").strip())
    if not path.is_absolute():
        path = root / path
    return DatabaseConfig(sqlite_path=path.resolve())


def init_db(db: str | Path | DatabaseConfig) -> None:
    config = _coerce_config(db)
    assert config.sqlite_path is not None
    config.sqlite_path.parent.mkdir(parents=True, exist_ok=True)
    with _connect(config) as conn:
        conn.executescript(SQLITE_SCHEMA)
        _ensure_question_columns(conn)
        _backfill_question_metadata(conn)
        _refresh_automatic_classification(conn)


def create_import_job(db, *, source_filename: str, stored_filename: str, preview_relpath: str, stats: dict) -> int:
    with _connect(_coerce_config(db)) as conn:
        cursor = conn.execute(
            "INSERT INTO imports (source_filename, stored_filename, preview_relpath, stats_json) VALUES (?, ?, ?, ?)",
            (source_filename, stored_filename, preview_relpath, json.dumps(stats, ensure_ascii=False)),
        )
        return int(cursor.lastrowid)


def save_questions(db, import_id: int, questions: Iterable[QuestionDraft]) -> None:
    rows = [(
        import_id, q.number, q.question_type, q.category_major, q.category_minor, q.difficulty,
        q.curriculum_version, q.curriculum_chapter, q.curriculum_section,
        q.classification_source, q.classification_confidence, q.review_reason,
        json.dumps(q.asset_manifest, ensure_ascii=False), q.stem, q.stem_html,
        "\n".join(q.options), q.options_html, q.answer, q.answer_html, q.analysis,
        q.analysis_html, q.source_name, q.raw_text, q.raw_html,
    ) for q in questions]
    if not rows:
        return
    with _connect(_coerce_config(db)) as conn:
        conn.executemany(
            """INSERT INTO questions
            (import_id, number, question_type, category_major, category_minor, difficulty,
             curriculum_version, curriculum_chapter, curriculum_section, classification_source,
             classification_confidence, review_reason, asset_manifest, stem, stem_html, options,
             options_html, answer, answer_html, analysis, analysis_html, source_name, raw_text, raw_html)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            rows,
        )


def list_questions(db, *, keyword="", question_type="", category_major="", category_minor="", difficulty="", import_id=None, limit=500):
    clauses, params = [], []
    if keyword:
        clauses.append("(q.stem LIKE ? OR q.answer LIKE ? OR q.analysis LIKE ? OR q.raw_text LIKE ?)")
        like = f"%{keyword}%"; params.extend([like] * 4)
    for column, value in (("question_type", question_type), ("category_minor", category_minor), ("difficulty", difficulty)):
        if value:
            operator, param = ("LIKE", f"%{value}%") if column == "category_minor" else ("=", value)
            clauses.append(f"q.{column} {operator} ?"); params.append(param)
    if category_major:
        if category_major == "未分类": clauses.append("(q.category_major IS NULL OR q.category_major = '')")
        else: clauses.append("q.category_major = ?"); params.append(category_major)
    if import_id is not None:
        clauses.append("q.import_id = ?"); params.append(import_id)
    where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    params.append(max(0, int(limit)))
    with _connect(_coerce_config(db)) as conn:
        rows = conn.execute(f"""SELECT q.*, i.preview_relpath, i.source_filename AS import_source
            FROM questions q JOIN imports i ON i.id=q.import_id {where}
            ORDER BY q.import_id DESC, CAST(q.number AS INTEGER), q.id LIMIT ?""", params).fetchall()
    return [_deserialize_question(row) for row in rows]


def get_questions_by_ids(db, ids: list[int] | None = None):
    with _connect(_coerce_config(db)) as conn:
        if ids is not None:
            unique = list(dict.fromkeys(int(v) for v in ids))
            if not unique:
                return []
            rows = conn.execute(f"SELECT * FROM questions WHERE id IN ({','.join('?' for _ in unique)})", unique).fetchall()
            by_id = {int(r["id"]): _deserialize_question(r) for r in rows}
            return [by_id[v] for v in unique if v in by_id]
        rows = conn.execute("SELECT * FROM questions ORDER BY import_id DESC, CAST(number AS INTEGER), id").fetchall()
    return [_deserialize_question(row) for row in rows]


def list_imports(db):
    with _connect(_coerce_config(db)) as conn:
        rows = conn.execute("""SELECT i.*, COALESCE(q.question_count,0) AS question_count FROM imports i
            LEFT JOIN (SELECT import_id,COUNT(*) AS question_count FROM questions GROUP BY import_id) q
            ON q.import_id=i.id ORDER BY i.id DESC""").fetchall()
    result = []
    for row in rows:
        item = dict(row); item["stats"] = _json_value(item.pop("stats_json"), {}); result.append(item)
    return result


def update_question(db, question_id: int, data: dict) -> None:
    allowed = {"number","question_type","category_major","category_minor","difficulty","curriculum_version",
        "curriculum_chapter","curriculum_section","classification_source","classification_confidence",
        "review_reason","asset_manifest","stem","stem_html","options","options_html","answer","answer_html",
        "analysis","analysis_html"}
    fields = [f for f in allowed if f in data]
    if not fields: return
    if any(field in data for field in ("question_type", "category_major", "category_minor", "difficulty", "curriculum_chapter", "curriculum_section")):
        data = dict(data)
        data.setdefault("classification_source", "manual")
        data.setdefault("classification_confidence", None)
        fields = [f for f in allowed if f in data]
    values = [json.dumps(data[f], ensure_ascii=False) if f == "asset_manifest" and not isinstance(data[f], str) else data[f] for f in fields]
    with _connect(_coerce_config(db)) as conn:
        conn.execute(f"UPDATE questions SET {', '.join(f'{f} = ?' for f in fields)} WHERE id = ?", [*values, question_id])


def delete_question(db, question_id: int) -> None:
    with _connect(_coerce_config(db)) as conn: conn.execute("DELETE FROM questions WHERE id=?", (question_id,))


def clear_library_data(db) -> None:
    with _connect(_coerce_config(db)) as conn:
        conn.execute("DELETE FROM papers"); conn.execute("DELETE FROM questions"); conn.execute("DELETE FROM imports")


def dashboard_stats(db):
    with _connect(_coerce_config(db)) as conn:
        qc = conn.execute("SELECT COUNT(*) count FROM questions").fetchone()["count"]
        ic = conn.execute("SELECT COUNT(*) count FROM imports").fetchone()["count"]
        types = conn.execute("SELECT question_type,COUNT(*) count FROM questions GROUP BY question_type ORDER BY count DESC").fetchall()
        cats = conn.execute("SELECT COALESCE(NULLIF(category_major,''),'未分类') category_major,COUNT(*) count FROM questions GROUP BY COALESCE(NULLIF(category_major,''),'未分类') ORDER BY count DESC").fetchall()
        diffs = conn.execute("SELECT COALESCE(NULLIF(difficulty,''),'未设置') difficulty,COUNT(*) count FROM questions GROUP BY COALESCE(NULLIF(difficulty,''),'未设置') ORDER BY count DESC").fetchall()
    return {"question_count":qc,"import_count":ic,"type_counts":[dict(r) for r in types],"category_counts":[dict(r) for r in cats],"difficulty_counts":[dict(r) for r in diffs]}


def create_paper(db, data: dict, question_ids: Sequence[int]) -> int:
    ids = _normalize_question_ids(question_ids)
    with _connect(_coerce_config(db)) as conn:
        values = tuple(
            _clean_paper_field(k, data.get(k), default=d)
            for k,d in (
                ("title","\u672a\u547d\u540d\u8bd5\u5377"),("exam_time",""),("total_score",""),("school",""),("grade",""),
                ("class_name",""),("instructions",""),("source","manual"),("status","saved"))
        )
        cursor = conn.execute("INSERT INTO papers (title,exam_time,total_score,school,grade,class_name,instructions,source,status) VALUES (?,?,?,?,?,?,?,?,?)", values)
        paper_id = int(cursor.lastrowid); _replace_paper_questions(conn, paper_id, ids); return paper_id

def update_paper(db, paper_id: int, data: dict, question_ids: Sequence[int] | None = None) -> None:
    allowed = {"title","exam_time","total_score","school","grade","class_name","instructions","source","status"}
    fields = [f for f in allowed if f in data]
    with _connect(_coerce_config(db)) as conn:
        if fields:
            values = [_clean_paper_field(f, data[f], paper_id=paper_id) for f in fields]
            conn.execute(f"UPDATE papers SET {', '.join(f'{f}=?' for f in fields)},updated_at=CURRENT_TIMESTAMP WHERE id=?", [*values, paper_id])
        if question_ids is not None:
            _replace_paper_questions(conn, paper_id, _normalize_question_ids(question_ids))
            conn.execute("UPDATE papers SET updated_at=CURRENT_TIMESTAMP WHERE id=?", (paper_id,))

def get_paper(db, paper_id: int):
    with _connect(_coerce_config(db)) as conn:
        paper = conn.execute("SELECT * FROM papers WHERE id=?", (paper_id,)).fetchone()
        if paper is None: return None
        rows = conn.execute("""SELECT pq.position,pq.display_number,pq.original_question_id,q.*,
            CASE WHEN q.id IS NULL THEN 1 ELSE 0 END missing FROM paper_questions pq
            LEFT JOIN questions q ON q.id=pq.question_id WHERE pq.paper_id=? ORDER BY pq.position,pq.id""", (paper_id,)).fetchall()
    result = _clean_paper_row(dict(paper)); result["questions"] = [_paper_question(r) for r in rows]; result["question_count"] = len(rows); return result


def list_papers(db, limit=200):
    with _connect(_coerce_config(db)) as conn:
        rows = conn.execute("""SELECT p.*,COUNT(pq.id) question_count FROM papers p LEFT JOIN paper_questions pq
            ON pq.paper_id=p.id GROUP BY p.id ORDER BY p.updated_at DESC,p.id DESC LIMIT ?""", (max(0,int(limit)),)).fetchall()
    return [_clean_paper_row(dict(r)) for r in rows]


def delete_paper(db, paper_id: int) -> None:
    with _connect(_coerce_config(db)) as conn: conn.execute("DELETE FROM papers WHERE id=?", (paper_id,))


def create_export_record(db, paper_id: int, export_format: str, filename: str, question_count: int) -> int:
    with _connect(_coerce_config(db)) as conn:
        cursor = conn.execute(
            "INSERT INTO export_records (paper_id, export_format, filename, question_count) VALUES (?, ?, ?, ?)",
            (paper_id, export_format, filename, int(question_count)),
        )
        return int(cursor.lastrowid)


def list_export_records(db, limit: int = 200) -> list[dict]:
    with _connect(_coerce_config(db)) as conn:
        rows = conn.execute(
            """SELECT e.*, p.title, p.exam_time, p.total_score
               FROM export_records e JOIN papers p ON p.id=e.paper_id
               ORDER BY e.created_at DESC, e.id DESC LIMIT ?""",
            (max(0, int(limit)),),
        ).fetchall()
    return [_clean_paper_row(dict(row)) for row in rows]


def storage_label(db) -> str: return f"SQLite: {_coerce_config(db).sqlite_path}"


def database_health(db) -> dict:
    config = _coerce_config(db)
    with _connect(config) as conn:
        return {"path":str(config.sqlite_path),"integrity":conn.execute("PRAGMA integrity_check").fetchone()[0],
            "journal_mode":str(conn.execute("PRAGMA journal_mode").fetchone()[0]).lower(),
            "foreign_keys":bool(conn.execute("PRAGMA foreign_keys").fetchone()[0])}


@contextmanager
def _connect(config: DatabaseConfig) -> Iterator[sqlite3.Connection]:
    if config.driver.lower() != "sqlite": raise ValueError("当前版本仅支持 SQLite 数据库。")
    if config.sqlite_path is None: raise ValueError("未配置 SQLite 数据库路径。")
    config.sqlite_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(config.sqlite_path, timeout=10.0); conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON"); conn.execute("PRAGMA busy_timeout=10000"); conn.execute("PRAGMA journal_mode=WAL")
    try:
        yield conn; conn.commit()
    except Exception:
        conn.rollback(); raise
    finally: conn.close()


def _ensure_question_columns(conn):
    existing = {r["name"] for r in conn.execute("PRAGMA table_info(questions)").fetchall()}
    for name, kind in QUESTION_COLUMNS.items():
        if name not in existing: conn.execute(f"ALTER TABLE questions ADD COLUMN {name} {kind}")


def _backfill_question_metadata(conn):
    rows = conn.execute("SELECT id,number,question_type,category_major,category_minor,difficulty,classification_source,stem,options,answer,analysis FROM questions").fetchall()
    for row in rows:
        data = dict(row); inferred = infer_question_metadata(data); updates = {}
        for field in ("question_type","category_major","category_minor","difficulty"):
            if not str(data.get(field) or "").strip(): updates[field] = inferred[field]
        if not str(data.get("classification_source") or "").strip(): updates["classification_source"] = "rule" if updates else "legacy"
        if updates: conn.execute(f"UPDATE questions SET {', '.join(f'{f}=?' for f in updates)} WHERE id=?", [*updates.values(),data["id"]])


def _refresh_automatic_classification(conn):
    rows = conn.execute(
        """SELECT id,number,question_type,category_major,category_minor,difficulty,stem,options,answer,analysis,
                  source_name,raw_text,stem_html,options_html,answer_html,analysis_html,raw_html,asset_manifest,
                  curriculum_version,curriculum_chapter,curriculum_section,classification_source,classification_confidence,review_reason
           FROM questions WHERE classification_source IN ('rule','standard_19')"""
    ).fetchall()
    for row in rows:
        data = dict(row)
        question = QuestionDraft(
            number=data.get("number") or "", stem=data.get("stem") or "", options=(data.get("options") or "").splitlines(),
            answer=data.get("answer") or "", analysis=data.get("analysis") or "", question_type=data.get("question_type") or "",
            source_name=data.get("source_name") or "", raw_text=data.get("raw_text") or "", stem_html=data.get("stem_html") or "",
            options_html=data.get("options_html") or "", answer_html=data.get("answer_html") or "", analysis_html=data.get("analysis_html") or "",
            raw_html=data.get("raw_html") or "", category_major="", category_minor="",
            difficulty=data.get("difficulty") or "", curriculum_version=data.get("curriculum_version") or "人教B版",
            curriculum_chapter=data.get("curriculum_chapter") or "", curriculum_section=data.get("curriculum_section") or "",
            classification_source=data.get("classification_source") or "rule", classification_confidence=data.get("classification_confidence"),
            review_reason=data.get("review_reason") or "", asset_manifest=_json_value(data.get("asset_manifest"), []),
        )
        updated = enrich_question_draft(question)
        if data.get("classification_source") == "standard_19":
            updated = replace(updated, question_type=question.question_type, difficulty=question.difficulty,
                              classification_source="standard_19", classification_confidence=1.0)
        conn.execute(
            """UPDATE questions SET question_type=?,category_major=?,category_minor=?,difficulty=?,curriculum_version=?,
                      curriculum_chapter=?,curriculum_section=?,classification_source=?,classification_confidence=?,review_reason=?
               WHERE id=?""",
            (updated.question_type, updated.category_major, updated.category_minor, updated.difficulty, updated.curriculum_version,
             updated.curriculum_chapter, updated.curriculum_section, updated.classification_source, updated.classification_confidence,
             updated.review_reason, data["id"]),
        )


def _replace_paper_questions(conn, paper_id: int, ids: list[int]):
    conn.execute("DELETE FROM paper_questions WHERE paper_id=?", (paper_id,))
    if not ids: return
    found = {int(r["id"]) for r in conn.execute(f"SELECT id FROM questions WHERE id IN ({','.join('?' for _ in ids)})", ids).fetchall()}
    conn.executemany("INSERT INTO paper_questions (paper_id,question_id,original_question_id,position,display_number) VALUES (?,?,?,?,?)",
        [(paper_id,qid,qid,pos,pos) for pos,qid in enumerate(ids,1) if qid in found])


def _normalize_question_ids(ids):
    values = list(dict.fromkeys(int(v) for v in ids if int(v)>0))
    if len(values)>100: raise ValueError("每份组卷最多包含 100 道题。")
    return values


def _deserialize_question(row):
    item = dict(row); item["asset_manifest"] = _json_value(item.get("asset_manifest"), []); return item


def _paper_question(row):
    item = dict(row); missing = bool(item.pop("missing")); original = int(item["original_question_id"])
    if missing: return {"id":None,"original_question_id":original,"position":int(item["position"]),"display_number":int(item["display_number"]),"missing":True}
    item["missing"] = False; item["asset_manifest"] = _json_value(item.get("asset_manifest"), []); return item


def _clean_paper_row(item: dict) -> dict:
    paper_id = item.get("paper_id") or item.get("id")
    for field in ("title", "exam_time", "total_score", "school", "grade", "class_name", "instructions"):
        if field in item:
            item[field] = _clean_paper_field(field, item.get(field), paper_id=paper_id)
    return item


def _clean_paper_field(field: str, value, *, default: str = "", paper_id=None) -> str:
    text = str(value if value is not None else default).strip()
    if field == "title":
        if not text or _looks_question_mark_garbled(text):
            return f"\u672a\u547d\u540d\u8bd5\u5377 #{paper_id}" if paper_id else "\u672a\u547d\u540d\u8bd5\u5377"
        return text
    if field == "exam_time":
        match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*[?\uff1f]+", text)
        if match:
            return f"{match.group(1)} \u5206\u949f"
        return "" if _looks_question_mark_garbled(text) else text
    if field == "total_score":
        match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*[?\uff1f]+", text)
        if match:
            return match.group(1)
    if field in {"school", "grade", "class_name", "instructions"} and _looks_question_mark_garbled(text):
        return ""
    return text


def _looks_question_mark_garbled(value: str) -> bool:
    text = str(value or "").strip()
    if not text:
        return False
    mark_count = sum(1 for ch in text if ch in "?\uff1f")
    meaningful_count = sum(1 for ch in text if not ch.isspace())
    return mark_count >= 2 and meaningful_count > 0 and mark_count / meaningful_count >= 0.7


def _json_value(value, fallback):
    if isinstance(value,(list,dict)): return value
    try: return json.loads(str(value or ""))
    except (TypeError,ValueError,json.JSONDecodeError): return fallback


def _coerce_config(db):
    if isinstance(db, DatabaseConfig):
        if db.driver.lower() != "sqlite": raise ValueError("当前版本仅支持 SQLite 数据库。")
        return db
    return DatabaseConfig(sqlite_path=Path(db).resolve())


def _read_dotenv(path: Path):
    if not path.exists(): return {}
    values = {}
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if line and not line.startswith("#") and "=" in line:
            key,value = line.split("=",1); values[key.strip()] = value.strip().strip('"').strip("'")
    return values
