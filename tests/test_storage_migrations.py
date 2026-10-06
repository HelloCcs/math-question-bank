from __future__ import annotations

import sqlite3

from question_bank.storage import database_health, init_db, list_questions, load_database_config


LEGACY_SCHEMA = """
CREATE TABLE imports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_filename TEXT NOT NULL,
    stored_filename TEXT NOT NULL,
    preview_relpath TEXT,
    stats_json TEXT NOT NULL DEFAULT '{}',
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE questions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_id INTEGER NOT NULL,
    number TEXT,
    question_type TEXT,
    category_major TEXT,
    category_minor TEXT,
    difficulty TEXT,
    stem TEXT NOT NULL,
    options TEXT,
    answer TEXT,
    analysis TEXT,
    source_name TEXT,
    raw_text TEXT,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
INSERT INTO imports (source_filename, stored_filename) VALUES ('legacy.docx', 'legacy-token.docx');
INSERT INTO questions
    (import_id, number, question_type, category_major, category_minor, difficulty,
     stem, options, answer, analysis, source_name, raw_text)
VALUES
    (1, '15', '解答题', '导数', '人工分类', '困难',
     '证明函数单调。', '', '略', '利用导数。', 'legacy.docx', '15. 证明函数单调。');
"""


def test_legacy_sqlite_migrates_without_overwriting_manual_values(tmp_path):
    path = tmp_path / "legacy.db"
    connection = sqlite3.connect(path)
    connection.executescript(LEGACY_SCHEMA)
    connection.commit()
    connection.close()

    init_db(path)
    row = list_questions(path)[0]
    assert row["question_type"] == "解答题"
    assert row["category_major"] == "导数"
    assert row["category_minor"] == "人工分类"
    assert row["difficulty"] == "困难"
    assert row["curriculum_version"] == "人教B版"
    assert row["asset_manifest"] == []

    columns = {item[1] for item in sqlite3.connect(path).execute("PRAGMA table_info(questions)")}
    assert {"classification_source", "classification_confidence", "review_reason", "asset_manifest"} <= columns
    tables = {item[0] for item in sqlite3.connect(path).execute("SELECT name FROM sqlite_master WHERE type='table'")}
    assert {"papers", "paper_questions"} <= tables


def test_sqlite_health_and_default_config(tmp_path, monkeypatch):
    monkeypatch.delenv("SQLITE_PATH", raising=False)
    config = load_database_config(tmp_path)
    assert config.driver == "sqlite"
    assert config.sqlite_path == (tmp_path / "data" / "question_bank.db").resolve()
    init_db(config)
    health = database_health(config)
    assert health["integrity"] == "ok"
    assert health["journal_mode"] == "wal"
    assert health["foreign_keys"] is True

