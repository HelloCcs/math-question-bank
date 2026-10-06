import json
from io import BytesIO
from question_bank.ai_classifier import (
    AIConfig,
    DEFAULT_AI_BASE_URL,
    DEFAULT_AI_MODEL,
    classify_import_questions,
    classify_low_confidence,
    load_ai_config,
    masked_ai_status,
    save_ai_config,
)
from question_bank.parser import QuestionDraft

def _q(): return QuestionDraft("1","函数单调性",[],"","","解答题","a.docx","1.函数单调性",classification_confidence=0.0)

def test_ai_disabled_and_incomplete_config():
    rows,warnings=classify_low_confidence([_q()],AIConfig())
    assert rows[0].classification_source=="rule" and warnings==[]
    rows,warnings=classify_low_confidence([_q()],AIConfig(enabled=True))
    assert warnings and "配置不完整" in warnings[0]

def test_ai_valid_json(monkeypatch):
    content=json.dumps({"results":[{"number":"1","category_major":"函数","category_minor":"函数性质","question_type":"解答题","difficulty":"中等","chapter":"函数","section":"函数性质","confidence":0.9}]},ensure_ascii=False)
    response={"choices":[{"message":{"content":content}}]}
    class Fake:
        def __enter__(self): return self
        def __exit__(self,*a): pass
        def read(self): return json.dumps(response,ensure_ascii=False).encode()
    monkeypatch.setattr("question_bank.ai_classifier.urlopen",lambda *a,**k:Fake())
    rows,warnings=classify_low_confidence([_q()],AIConfig(True,"http://local","secret","model"))
    assert warnings==[] and rows[0].classification_source=="ai" and rows[0].curriculum_chapter=="函数"

def test_ai_without_valid_major_category_is_not_marked_ai(monkeypatch):
    content=json.dumps({"results":[{"number":"1","question_type":"解答题","difficulty":"中等","confidence":1}]},ensure_ascii=False)
    response={"choices":[{"message":{"content":content}}]}
    class Fake:
        def __enter__(self): return self
        def __exit__(self,*a): pass
        def read(self): return json.dumps(response,ensure_ascii=False).encode()
    monkeypatch.setattr("question_bank.ai_classifier.urlopen",lambda *a,**k:Fake())
    rows,warnings=classify_import_questions([_q()],AIConfig(True,"http://local","secret","model"))
    assert rows[0].classification_source=="rule"
    assert warnings and "没有返回合法大分类" in warnings[0]

def test_ai_reclassifies_all_import_questions(monkeypatch):
    q1 = _q()
    q2 = QuestionDraft(
        "2","sin 和 cos 的函数综合题",[],"","","解答题","a.docx","2.sin 和 cos",
        category_major="三角函数", category_minor="三角恒等变换", classification_confidence=0.9,
    )
    content=json.dumps({"results":[
        {"number":"1","category_major":"导数","category_minor":"导数应用","question_type":"解答题","difficulty":"较难","confidence":0.96},
        {"number":"2","category_major":"三角函数","category_minor":"三角恒等变换","question_type":"解答题","difficulty":"中等","confidence":0.92},
    ]},ensure_ascii=False)
    response={"choices":[{"message":{"content":content}}]}
    class Fake:
        def __enter__(self): return self
        def __exit__(self,*a): pass
        def read(self): return json.dumps(response,ensure_ascii=False).encode()
    monkeypatch.setattr("question_bank.ai_classifier.urlopen",lambda *a,**k:Fake())
    rows,warnings=classify_import_questions([q1,q2],AIConfig(True,"http://local","secret","model"))
    assert warnings==[]
    assert rows[0].classification_source=="ai" and rows[0].category_major=="导数"
    assert rows[1].classification_source=="ai" and rows[1].category_major=="三角函数"
    assert rows[0].rule_category_major != rows[0].ai_category_major
    assert rows[0].ai_review_note.startswith("AI改分类")
    assert rows[1].ai_review_note.startswith("AI确认本地分类")

def test_ai_missing_confidence_does_not_become_zero(monkeypatch):
    content=json.dumps({"results":[
        {"number":"1","category_major":"导数","category_minor":"导数应用","question_type":"解答题","difficulty":"中等"},
    ]},ensure_ascii=False)
    response={"choices":[{"message":{"content":content}}]}
    class Fake:
        def __enter__(self): return self
        def __exit__(self,*a): pass
        def read(self): return json.dumps(response,ensure_ascii=False).encode()
    monkeypatch.setattr("question_bank.ai_classifier.urlopen",lambda *a,**k:Fake())
    rows,warnings=classify_import_questions([_q()],AIConfig(True,"http://local","secret","model"))
    assert warnings==[]
    assert rows[0].classification_source=="ai"
    assert rows[0].classification_confidence is None
    assert rows[0].review_reason==""

def test_ai_accepts_percent_confidence(monkeypatch):
    content=json.dumps({"results":[
        {"number":"1","category_major":"导数","category_minor":"导数应用","question_type":"解答题","difficulty":"中等","confidence":92},
    ]},ensure_ascii=False)
    response={"choices":[{"message":{"content":content}}]}
    class Fake:
        def __enter__(self): return self
        def __exit__(self,*a): pass
        def read(self): return json.dumps(response,ensure_ascii=False).encode()
    monkeypatch.setattr("question_bank.ai_classifier.urlopen",lambda *a,**k:Fake())
    rows,warnings=classify_import_questions([_q()],AIConfig(True,"http://local","secret","model"))
    assert warnings==[]
    assert rows[0].classification_confidence==0.92

def test_ai_invalid_json_falls_back(monkeypatch):
    class Fake:
        def __enter__(self): return self
        def __exit__(self,*a): pass
        def read(self): return b'not json'
    monkeypatch.setattr("question_bank.ai_classifier.urlopen",lambda *a,**k:Fake())
    rows,warnings=classify_low_confidence([_q()],AIConfig(True,"http://local","secret","model"))
    assert rows[0].classification_source=="rule" and warnings

def test_masked_status_never_exposes_full_key():
    status=masked_ai_status(AIConfig(True,"http://x","super-secret-1234","m"))
    assert status["api_key"]=="***1234" and "super-secret" not in str(status)

def test_load_config_from_mapping():
    config=load_ai_config(env={"AI_ENABLED":"true","AI_BASE_URL":"http://x/","AI_API_KEY":"k","AI_MODEL":"m"})
    assert config.enabled and config.base_url=="http://x"

def test_api_key_only_uses_default_endpoint_and_model():
    config=load_ai_config(env={"AI_API_KEY":"sk-test"})
    assert config.enabled
    assert config.base_url==DEFAULT_AI_BASE_URL
    assert config.model==DEFAULT_AI_MODEL

def test_save_api_key_only_writes_default_endpoint_and_model(tmp_path):
    save_ai_config(tmp_path, {"api_key":"sk-test"})
    saved=load_ai_config(tmp_path)
    assert saved.enabled
    assert saved.api_key=="sk-test"
    assert saved.base_url==DEFAULT_AI_BASE_URL
    assert saved.model==DEFAULT_AI_MODEL
