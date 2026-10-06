from question_bank.curriculum import chapter_names, classify_curriculum, load_curriculum

def test_curriculum_config_is_valid():
    data = load_curriculum()
    assert data["version"] == "人教B版"
    assert "函数" in chapter_names()
    assert all(chapter["sections"] for chapter in data["chapters"])

def test_curriculum_classification():
    result = classify_curriculum("求函数的定义域并判断单调性")
    assert result["chapter"] == "函数"
    assert result["section"] == "函数性质"
    assert result["confidence"] > 0.5

def test_strong_geometry_and_conic_clues_beat_generic_words():
    geometry = classify_curriculum("四棱锥，平面垂直，二面角，求体积和外接球表面积")
    conic = classify_curriculum("曲线C上的动点M满足MF2-MF1=4，F1、F2为焦点，求C的方程和取值范围")
    assert geometry["chapter"] == "立体几何"
    assert geometry["section"] == "空间几何"
    assert conic["chapter"] == "解析几何"
    assert conic["section"] == "圆锥曲线"
