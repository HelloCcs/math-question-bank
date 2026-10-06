from __future__ import annotations

import pytest

from question_bank.parser import QuestionDraft
from question_bank.storage import (
    create_import_job,
    create_export_record,
    create_paper,
    delete_paper,
    delete_question,
    get_paper,
    get_questions_by_ids,
    init_db,
    list_export_records,
    list_papers,
    list_questions,
    save_questions,
    update_paper,
)


def _seed_many(db_config, count=3):
    init_db(db_config)
    import_id = create_import_job(
        db_config,
        source_filename="paper.docx",
        stored_filename="paper-token.docx",
        preview_relpath="previews/paper/preview.htm",
        stats={"question_count": count},
    )
    save_questions(
        db_config,
        import_id,
        [
            QuestionDraft(
                number=str(index), stem=f"题目 {index}", options=[], answer=str(index), analysis="解析",
                question_type="解答题", source_name="paper.docx", raw_text=f"{index}. 题目 {index}",
            )
            for index in range(1, count + 1)
        ],
    )
    return [row["id"] for row in list_questions(db_config)]


def test_paper_crud_order_and_missing_question(db_config):
    ids = _seed_many(db_config)
    requested = [ids[2], ids[0], ids[2], ids[1]]
    paper_id = create_paper(
        db_config,
        {"title": "阶段一试卷", "exam_time": "120 分钟", "total_score": "150", "source": "manual"},
        requested,
    )
    paper = get_paper(db_config, paper_id)
    assert paper["title"] == "阶段一试卷"
    assert [row["id"] for row in paper["questions"]] == [ids[2], ids[0], ids[1]]
    assert [row["display_number"] for row in paper["questions"]] == [1, 2, 3]
    assert list_papers(db_config)[0]["question_count"] == 3

    update_paper(db_config, paper_id, {"title": "已调整试卷"}, [ids[1], ids[2]])
    paper = get_paper(db_config, paper_id)
    assert paper["title"] == "已调整试卷"
    assert [row["id"] for row in paper["questions"]] == [ids[1], ids[2]]

    delete_question(db_config, ids[1])
    paper = get_paper(db_config, paper_id)
    assert paper["questions"][0]["missing"] is True
    assert paper["questions"][0]["original_question_id"] == ids[1]

    delete_paper(db_config, paper_id)
    assert get_paper(db_config, paper_id) is None


def test_paper_question_mark_garble_is_cleaned(db_config):
    ids = _seed_many(db_config, 1)
    paper_id = create_paper(
        db_config,
        {"title": "????????", "exam_time": "120 ??", "total_score": "150 ??", "school": "??"},
        ids,
    )
    create_export_record(db_config, paper_id, "Word ZIP", "paper.zip", 1)

    paper = get_paper(db_config, paper_id)
    assert paper["title"] == "\u672a\u547d\u540d\u8bd5\u5377"
    assert paper["exam_time"] == "120 \u5206\u949f"
    assert paper["total_score"] == "150"
    assert paper["school"] == ""

    listed = list_papers(db_config)[0]
    assert listed["title"] == "\u672a\u547d\u540d\u8bd5\u5377"
    assert listed["exam_time"] == "120 \u5206\u949f"

    exported = list_export_records(db_config)[0]
    assert exported["title"] == "\u672a\u547d\u540d\u8bd5\u5377"
    assert exported["exam_time"] == "120 \u5206\u949f"


def test_get_questions_preserves_requested_order(db_config):
    ids = _seed_many(db_config)
    assert [row["id"] for row in get_questions_by_ids(db_config, [ids[2], ids[0]])] == [ids[2], ids[0]]


def test_paper_rejects_more_than_100_questions(db_config):
    ids = _seed_many(db_config, 101)
    with pytest.raises(ValueError, match="100"):
        create_paper(db_config, {"title": "超限"}, ids)
