from __future__ import annotations
import json
from functools import lru_cache
from pathlib import Path

DATA_PATH = Path(__file__).parent / "data" / "pep_b_curriculum.json"

@lru_cache(maxsize=1)
def load_curriculum() -> dict:
    data = json.loads(DATA_PATH.read_text(encoding="utf-8"))
    if not data.get("version") or not isinstance(data.get("chapters"), list):
        raise ValueError("人教B版目录配置无效。")
    return data

def classify_curriculum(text: str) -> dict:
    content = (text or "").lower()
    # Strong structural clues must beat generic words such as "三角形" or
    # "取值范围", which often appear in unrelated geometry and conic problems.
    strong_rules = [
        ("立体几何", "空间几何", ("四棱锥", "三棱锥", "四棱柱", "三棱柱", "棱锥", "棱柱", "二面角", "外接球", "截面"), 8),
        ("解析几何", "圆锥曲线", ("双曲线", "椭圆", "抛物线", "焦点弦", "离心率", "渐近线", "mf2", "mf1"), 8),
        ("概率统计", "概率统计", ("概率", "随机变量", "分布列", "数学期望", "条件概率", "随机抽取", "独立事件"), 8),
        ("导数及其应用", "导数应用", ("导数", "导函数", "求导", "切线斜率", "极值"), 8),
        ("复数", "复数", ("复数", "虚数", "共轭", "辐角"), 8),
        ("数列", "数列", ("数列", "等差", "等比", "通项公式", "递推"), 8),
        ("三角函数", "三角函数", ("sin", "cos", "tan", "正弦函数", "余弦函数", "正切函数", "三角恒等变换"), 8),
        ("三角函数", "解三角形", ("正弦定理", "余弦定理", "解三角形", "外接圆半径"), 8),
    ]
    strong_best = None
    for chapter, section, words, weight in strong_rules:
        hits = [word for word in words if word.lower() in content]
        if hits and (strong_best is None or (len(hits) * weight, len(hits)) > (strong_best[0], strong_best[1])):
            strong_best = (len(hits) * weight, len(hits), chapter, section, hits)
    if strong_best:
        return {"version": "人教B版", "chapter": strong_best[2], "section": strong_best[3],
                "confidence": min(0.98, 0.78 + 0.05 * strong_best[1])}

    best = None
    for chapter in load_curriculum()["chapters"]:
        for section in chapter["sections"]:
            words = [word for word in section.get("keywords", []) if word not in {"三角形", "边长", "面积", "最值", "取值范围", "直线", "平面", "方程"}]
            hits = [word for word in words if word.lower() in content]
            if hits and (best is None or len(hits) > best[0]):
                best = (len(hits), chapter["name"], section["name"], hits)
    if best is None:
        return {"version":"人教B版","chapter":"","section":"","confidence":0.0}
    return {"version":"人教B版","chapter":best[1],"section":best[2],"confidence":min(0.95,0.55+0.1*best[0])}

def chapter_names() -> list[str]:
    return [item["name"] for item in load_curriculum()["chapters"]]
