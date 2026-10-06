from __future__ import annotations

from pathlib import Path

import pytest

from question_bank.app import create_app
from question_bank.parser import QuestionDraft
from question_bank.storage import DatabaseConfig, create_import_job, save_questions


@pytest.fixture()
def db_config(tmp_path: Path) -> DatabaseConfig:
    return DatabaseConfig(driver="sqlite", sqlite_path=tmp_path / "question_bank.db")


@pytest.fixture()
def app(tmp_path: Path, db_config: DatabaseConfig):
    data_root = tmp_path / "data"
    config = {
        "TESTING": True,
        "SECRET_KEY": "stage-zero-tests",
        "DATA_ROOT": data_root,
        "DB_CONFIG": db_config,
        "ORIGINAL_DIR": data_root / "originals",
        "PREVIEW_DIR": data_root / "previews",
        "EXPORT_DIR": data_root / "exports",
        "PENDING_DIR": data_root / "pending",
    }
    return create_app(config)


@pytest.fixture()
def client(app):
    return app.test_client()


@pytest.fixture()
def seeded_question(db_config: DatabaseConfig) -> int:
    import_id = create_import_job(
        db_config,
        source_filename="baseline.docx",
        stored_filename="baseline-token.docx",
        preview_relpath="previews/baseline/preview.htm",
        stats={"question_count": 1, "import_mode": "word"},
    )
    save_questions(
        db_config,
        import_id,
        [
            QuestionDraft(
                number="1",
                stem="已知集合 A，求其补集。",
                options=["A. 空集", "B. 全集", "C. A", "D. 无法确定"],
                answer="B",
                analysis="根据补集定义可得。",
                question_type="单选题",
                source_name="baseline.docx",
                raw_text="1. 已知集合 A，求其补集。",
                stem_html="<p>已知集合 A，求其补集。</p>",
                options_html="<p>A. 空集</p><p>B. 全集</p><p>C. A</p><p>D. 无法确定</p>",
                answer_html="<p>B</p>",
                analysis_html="<p>根据补集定义可得。</p>",
                category_major="集合",
                category_minor="集合运算",
                difficulty="简单",
            )
        ],
    )
    return import_id

