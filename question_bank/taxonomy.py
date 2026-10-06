from __future__ import annotations

import re
from dataclasses import replace
from typing import Mapping

from .parser import QuestionDraft
from .curriculum import classify_curriculum


MAJOR_CATEGORIES = [
    "集合",
    "不等式",
    "函数",
    "导数",
    "三角函数",
    "解三角形",
    "数列",
    "空间几何",
    "解析几何",
    "概率统计",
    "向量",
]

DIFFICULTIES = ["简单", "中等", "较难", "困难", "挑战"]

QUESTION_TYPES = ["单选题", "多选题", "填空题", "解答题", "未分类", "OCR识别"]

_PRIORITY_CATEGORY_KEYWORDS: list[tuple[str, list[tuple[str, str]]]] = [
    ("导数", [("导数应用", r"导数|导函数|求导|切线方程|切线|极值|零点|驻点|单调区间|微分")]),
    ("解析几何", [
        ("解析几何", r"解析几何|圆锥曲线综合|圆锥曲线"),
        ("椭圆", r"椭圆|焦点|焦距|离心率|长轴|短轴|准线|标准方程"),
        ("双曲线", r"双曲线|实轴|虚轴|渐近线|离心率|焦点|焦距"),
        ("抛物线", r"抛物线|焦点|准线|顶点|对称轴"),
        ("圆锥曲线综合", r"圆锥曲线|曲线|直线与圆|圆心|半径|斜率|交点|位置关系|焦点弦|参数方程|极坐标"),
    ]),
    ("概率统计", [("概率统计", r"概率|统计|随机|分布|分布列|离散|方差|期望|数据|抽样|回归|独立性|样本|频率")]),
    ("空间几何", [("立体几何", r"空间|平面|四面体|棱柱|棱锥|圆柱|圆锥|几何体|二面角|体积|表面积")]),
    ("解三角形", [("解三角形", r"解三角形|正弦定理|余弦定理|三角形|△ABC|△|▽|角A|对边|内角|边a|边b|边c|a,b,c|A,B,C|外接圆|内切圆")]),
    ("三角函数", [("三角恒等变换", r"sin|cos|tan|正弦|余弦|正切|三角函数|周期|振幅|相位|三角恒等变换")]),
    ("集合", [("集合运算", r"集合|交集|并集|补集|子集|元素|空集|全集|属于|包含|充分不必要条件|充分必要条件|[∩∪]")]),
    ("不等式", [("不等式求解", r"不等式|恒成立|取值范围|放缩|均值不等式|柯西|已知实数|[<>≤≥]")]),
    ("数列", [("数列通项", r"数列|通项|前n项|前 n 项|项和|求和|递推|等差|等比")]),
    ("向量", [("向量运算", r"向量|数量积|坐标表示|基底|共线|垂直|模长|⃗|→")]),
    ("函数", [
        ("复数", r"复数|虚数|虚数单位|虚部|实部|共轭|辐角|z\b|\bi\b"),
        ("函数性质", r"函数|f\(x\)|log|奇函数|偶函数|对称中心|对称轴|定义域|值域|单调|零点|图象|图像|周期"),
        ("指数对数", r"指数函数|对数函数|幂函数|底数|真数|换底公式"),
    ]),
]

_CATEGORY_KEYWORDS: list[tuple[str, list[tuple[str, str]]]] = [
    ("集合", [("集合运算", "集合|交集|并集|补集|子集|元素|空集|全集|属于|包含")]),
    ("不等式", [("不等式求解", "不等式|恒成立|取值范围|最值|放缩|均值不等式|柯西")]),
    ("导数", [("导数应用", "导数|导函数|切线|极值|单调区间|最值|求导|微分")]),
    ("三角函数", [("三角恒等变换", "sin|cos|tan|正弦|余弦|正切|三角函数|周期|振幅|相位")]),
    ("解三角形", [("解三角形", "正弦定理|余弦定理|三角形|边长|面积|外接圆|内切圆")]),
    ("数列", [("数列通项", "数列|等差|等比|通项|前n项|前 n 项|递推|求和")]),
    ("空间几何", [("立体几何", "空间|平面|直线|四面体|棱锥|棱柱|垂直|二面角|体积|表面积")]),
    ("解析几何", [("椭圆", "椭圆|焦点|焦距|离心率|长轴|短轴|标准方程|准线")]),
    ("解析几何", [("双曲线", "双曲线|焦点|焦距|离心率|实轴|虚轴|渐近线|标准方程")]),
    ("解析几何", [("抛物线", "抛物线|焦点|准线|对称轴|顶点|标准方程")]),
    ("解析几何", [("直线与圆", "直线|圆|圆心|半径|方程|斜率|交点|位置关系")]),
    ("解析几何", [("圆锥曲线综合", "圆锥曲线|焦点弦|切线方程|参数方程|极坐标")]),
    ("概率统计", [("概率统计", "概率|统计|随机|分布|期望|方差|频率|样本|抽样")]),
    ("向量", [("向量运算", "向量|数量积|坐标表示|基底|共线|垂直|模长")]),
    ("函数", [("复数", "复数|虚数单位|虚部|实部|共轭|模|辐角")]),
    ("函数", [("函数性质", "函数|定义域|值域|单调|奇偶|零点|图象|图像|f\\(x\\)|周期")]),
    ("函数", [("指数对数", "指数函数|对数函数|幂函数|底数|真数|换底公式")]),
    ("函数", [("函数综合", "分段函数|复合函数|反函数|函数方程")]),
    ("概率统计", [("排列组合", "排列|组合|二项式|概率公式|计数原理")]),
]


def enrich_question_draft(question: QuestionDraft) -> QuestionDraft:
    metadata = infer_question_metadata(
        {
            "number": question.number,
            "question_type": question.question_type,
            "stem": question.stem,
            "options": "\n".join(question.options),
            "answer": question.answer,
            "analysis": question.analysis,
            "category_major": question.category_major,
            "category_minor": question.category_minor,
            "difficulty": question.difficulty,
        }
    )
    primary_text = _combined_text({"stem": question.stem, "options": "\n".join(question.options), "answer": question.answer})
    curriculum = classify_curriculum(primary_text)
    if not curriculum["chapter"]:
        curriculum = classify_curriculum(primary_text + "\n" + question.analysis)
    category_major = metadata["category_major"]
    category_minor = metadata["category_minor"]
    major_by_chapter = {
        "集合与常用逻辑用语": "集合", "等式与不等式": "不等式", "函数": "函数",
        "三角函数": "三角函数", "平面向量": "向量", "数列": "数列",
        "立体几何": "空间几何", "解析几何": "解析几何", "概率与统计": "概率统计", "概率统计": "概率统计",
        "导数及其应用": "导数", "复数": "函数",
    }
    curriculum_major = major_by_chapter.get(curriculum["chapter"], "")
    if curriculum_major and not category_major:
        category_major = curriculum_major
        if not category_minor or question.classification_source == "standard_19":
            category_minor = curriculum["section"]
    elif curriculum_major and question.classification_source == "standard_19":
        category_major = curriculum_major
        category_minor = curriculum["section"]
    return replace(
        question,
        question_type=metadata["question_type"],
        category_major=category_major,
        category_minor=category_minor,
        difficulty=metadata["difficulty"],
        curriculum_version=curriculum["version"],
        curriculum_chapter=curriculum["chapter"],
        curriculum_section=curriculum["section"],
        classification_source="rule",
        classification_confidence=curriculum["confidence"],
        review_reason="" if curriculum["chapter"] else "教材章节待人工确认",
    )


def enrich_questions(questions: list[QuestionDraft]) -> list[QuestionDraft]:
    enriched = [enrich_question_draft(question) for question in questions]
    numbers = [_question_number(question.number) for question in questions]
    is_standard_19 = len(questions) == 19 and numbers == list(range(1, 20))
    if not is_standard_19:
        return enriched
    result = []
    for question, number in zip(enriched, numbers):
        if number <= 8: question_type = "单选题"
        elif number <= 11: question_type = "多选题"
        elif number <= 14: question_type = "填空题"
        else: question_type = "解答题"
        if number in {1,2,3,4,5,9,15,16}: difficulty = "简单"
        elif number in {6,7,10,12,13,17}: difficulty = "中等"
        elif number in {8,11,14,18}: difficulty = "困难"
        else: difficulty = "挑战"
        result.append(replace(question, question_type=question_type, difficulty=difficulty,
            classification_source="standard_19", classification_confidence=1.0))
    return result


def infer_question_metadata(question: Mapping[str, object]) -> dict[str, str]:
    text = _combined_text(question)
    inferred_type = _infer_question_type(question)
    question_type = inferred_type or _clean(question.get("question_type")) or "解答题"

    category_major = _clean(question.get("category_major"))
    category_minor = _clean(question.get("category_minor"))
    if category_major not in MAJOR_CATEGORIES:
        category_major = ""
        category_minor = ""
    if not category_major:
        category_major, category_minor = _infer_category(text)
    if not category_major:
        category_major, category_minor = "函数", "待人工确认"
    elif not category_minor:
        _, category_minor = _infer_category(text, preferred_major=category_major)
        category_minor = category_minor or "待人工确认"

    difficulty = _clean(question.get("difficulty")) or _infer_difficulty(question)
    if difficulty not in DIFFICULTIES:
        difficulty = _infer_difficulty(question)

    return {
        "question_type": question_type or "未分类",
        "category_major": category_major,
        "category_minor": category_minor,
        "difficulty": difficulty,
    }


def _combined_text(question: Mapping[str, object]) -> str:
    parts = [
        _clean(question.get("stem")),
        _clean(question.get("options")),
        _clean(question.get("answer")),
        _clean(question.get("analysis")),
    ]
    return "\n".join(part for part in parts if part)


def _infer_question_type(question: Mapping[str, object]) -> str:
    number = _question_number(question.get("number"))
    answer = _clean(question.get("answer"))
    options = _clean(question.get("options"))
    if number:
        if 1 <= number <= 8:
            return "单选题"
        if 9 <= number <= 11:
            return "多选题"
        if 12 <= number <= 15:
            return "填空题"
        if number >= 16:
            return "解答题"
    if options:
        if re.fullmatch(r"[A-H](?:\s*[,，、]\s*[A-H])+", answer, re.IGNORECASE):
            return "多选题"
        return "单选题"
    if answer and len(answer) <= 12:
        return "填空题"
    return "解答题"


def _infer_category(text: str, preferred_major: str = "") -> tuple[str, str]:
    buckets = _PRIORITY_CATEGORY_KEYWORDS
    if preferred_major:
        preferred_priority = [item for item in _PRIORITY_CATEGORY_KEYWORDS if item[0] == preferred_major]
        if preferred_priority:
            buckets = preferred_priority
    for major, patterns in buckets:
        for minor, pattern in patterns:
            if re.search(pattern, text, flags=re.IGNORECASE):
                return major, minor

    buckets = _CATEGORY_KEYWORDS
    if preferred_major:
        buckets = [item for item in _CATEGORY_KEYWORDS if item[0] == preferred_major] or buckets
    for major, patterns in buckets:
        for minor, pattern in patterns:
            if re.search(pattern, text, flags=re.IGNORECASE):
                return major, minor
    return "", ""


def _infer_difficulty(question: Mapping[str, object]) -> str:
    number = _question_number(question.get("number"))
    if number:
        if number <= 8:
            return "简单"
        if number <= 13:
            return "中等"
        if number <= 17:
            return "较难"
        return "困难"
    text = _combined_text(question)
    if any(word in text for word in ("证明", "综合", "最值", "参数")):
        return "较难"
    return "中等"


def _question_number(value: object) -> int | None:
    text = _clean(value)
    return int(text) if text.isdigit() else None


def _clean(value: object) -> str:
    return str(value or "").strip()
