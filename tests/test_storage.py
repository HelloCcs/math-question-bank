from __future__ import annotations

from question_bank.parser import QuestionDraft
from question_bank.storage import (
    clear_library_data,
    create_import_job,
    dashboard_stats,
    delete_question,
    get_questions_by_ids,
    init_db,
    list_imports,
    list_questions,
    save_questions,
    storage_label,
    update_question,
)


def test_sqlite_storage_crud(db_config):
    init_db(db_config)
    import_id = create_import_job(
        db_config,
        source_filename="sample.docx",
        stored_filename="sample-token.docx",
        preview_relpath="previews/sample/preview.htm",
        stats={"question_count": 1},
    )
    save_questions(
        db_config,
        import_id,
        [
            QuestionDraft(
                number="15",
                stem="证明函数在区间内单调。",
                options=[],
                answer="略",
                analysis="利用导数证明。",
                question_type="解答题",
                source_name="sample.docx",
                raw_text="15. 证明函数在区间内单调。",
                category_major="导数",
                category_minor="导数应用",
                difficulty="较难",
            )
        ],
    )

    rows = list_questions(db_config, keyword="单调", question_type="解答题")
    assert len(rows) == 1
    assert rows[0]["number"] == "15"
    assert list_imports(db_config)[0]["question_count"] == 1
    assert dashboard_stats(db_config)["question_count"] == 1
    assert storage_label(db_config).startswith("SQLite:")

    question_id = rows[0]["id"]
    update_question(db_config, question_id, {"difficulty": "困难", "answer": "证明完成"})
    updated = get_questions_by_ids(db_config, [question_id])[0]
    assert updated["difficulty"] == "困难"
    assert updated["answer"] == "证明完成"

    delete_question(db_config, question_id)
    assert list_questions(db_config) == []
    clear_library_data(db_config)
    assert list_imports(db_config) == []


def test_init_db_is_idempotent(db_config):
    init_db(db_config)
    init_db(db_config)
    assert dashboard_stats(db_config) == {
        "question_count": 0,
        "import_count": 0,
        "type_counts": [],
        "category_counts": [],
        "difficulty_counts": [],
    }

