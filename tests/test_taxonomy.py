from dataclasses import replace
from question_bank.parser import QuestionDraft
from question_bank.taxonomy import enrich_question_draft, enrich_questions

def _question(number: int) -> QuestionDraft:
    return QuestionDraft(str(number), f"第 {number} 题函数题", [], "", "", "", "test.docx", f"{number}. 函数题")

def test_standard_19_question_types_and_difficulties():
    rows = enrich_questions([_question(i) for i in range(1, 20)])
    expected_types = ["单选题"]*8 + ["多选题"]*3 + ["填空题"]*3 + ["解答题"]*5
    expected_difficulties = {
        **{i:"简单" for i in [1,2,3,4,5,9,15,16]},
        **{i:"中等" for i in [6,7,10,12,13,17]},
        **{i:"困难" for i in [8,11,14,18]}, 19:"挑战",
    }
    assert [row.question_type for row in rows] == expected_types
    assert [row.difficulty for row in rows] == [expected_difficulties[i] for i in range(1,20)]
    assert all(row.classification_source == "standard_19" for row in rows)

def test_non_19_paper_does_not_force_number_rule():
    rows = enrich_questions([_question(i) for i in range(1, 16)])
    assert rows[14].classification_source == "rule"
    assert rows[14].question_type != "解答题" or rows[14].classification_confidence != 1.0

def test_existing_manual_category_is_preserved_by_storage_inference():
    row = enrich_question_draft(replace(_question(1), category_major="导数", category_minor="人工章节"))
    assert row.category_major == "导数"
    assert row.category_minor == "人工章节"

def test_priority_major_category_rules_choose_more_specific_topic():
    trig = enrich_question_draft(replace(_question(1), stem="已知函数 f(x)=sin x，求三角函数性质", answer=""))
    assert trig.category_major == "三角函数"

    deriv = enrich_question_draft(replace(_question(1), stem="已知函数与三角函数复合，讨论导数极值", answer=""))
    assert deriv.category_major == "导数"

    vector = enrich_question_draft(replace(_question(1), stem="解析几何与向量结合求解", answer=""))
    assert vector.category_major == "解析几何"
