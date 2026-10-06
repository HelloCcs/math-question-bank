from __future__ import annotations

import io
import json

from question_bank.storage import list_questions
from question_bank.app import _draft_from_dict


def test_primary_pages_render(client):
    expectations = {
        "/": "高中数学题库",
        "/import": "导入文件",
        "/manual": "单题导入",
        "/questions": "题库管理",
        "/paper": "智能组卷",
        "/settings/ai": "AI设置",
    }
    for path, marker in expectations.items():
        response = client.get(path)
        assert response.status_code == 200
        assert marker.encode("utf-8") in response.data


def test_question_update_and_delete(client, app, seeded_question):
    db = app.config["DB_CONFIG"]
    question_id = list_questions(db)[0]["id"]
    response = client.post(
        f"/questions/{question_id}/update",
        data={"difficulty": "中等", "number": "2"},
        follow_redirects=True,
    )
    assert response.status_code == 200
    updated = list_questions(db)[0]
    assert updated["difficulty"] == "中等"
    assert updated["number"] == "2"

    response = client.post(f"/questions/{question_id}/delete", follow_redirects=True)
    assert response.status_code == 200
    assert list_questions(db) == []


def test_empty_export_redirects_with_message(client):
    response = client.get("/export/xlsx", follow_redirects=True)
    assert response.status_code == 200
    assert "没有可导出的题目".encode("utf-8") in response.data


def test_invalid_upload_is_rejected(client):
    response = client.post(
        "/import",
        data={"file": (io.BytesIO(b"not a document"), "sample.txt")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )
    assert response.status_code == 200
    assert "当前版本支持".encode("utf-8") in response.data


def test_confirm_import_shows_ai_reclassification_status(client, app):
    token = "ai-status"
    payload = {
        "token": token,
        "source_filename": "sample.docx",
        "stored_filename": "sample.docx",
        "preview_relpath": "previews/sample/preview.html",
        "stats": {"question_count": 1, "import_mode": "word"},
        "questions": [
            {
                "number": "1",
                "stem": "求导函数单调性",
                "options": [],
                "answer": "",
                "analysis": "",
                "question_type": "解答题",
                "source_name": "sample.docx",
                "raw_text": "求导函数单调性",
                "category_major": "导数",
                "category_minor": "导数应用",
                "difficulty": "中等",
                "classification_source": "ai",
                "classification_confidence": 0.91,
                "rule_category_major": "函数",
                "rule_category_minor": "函数性质",
                "ai_category_major": "导数",
                "ai_category_minor": "导数应用",
                "ai_review_note": "AI改分类：规则函数 -> AI导数",
            }
        ],
    }
    path = app.config["PENDING_DIR"] / f"{token}.json"
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

    response = client.get(f"/import/confirm/{token}")

    assert response.status_code == 200
    assert "AI 复识别完成：1/1 题已由 AI 判断".encode("utf-8") in response.data
    assert "AI改动 1 题".encode("utf-8") in response.data
    assert "AI已复识别".encode("utf-8") in response.data
    assert "AI改分类：规则函数 -&gt; AI导数".encode("utf-8") in response.data


def test_confirmation_draft_preserves_extended_metadata():
    draft = _draft_from_dict({
        "number": "19", "stem": "题目", "options": [], "answer": "", "analysis": "解析",
        "question_type": "解答题", "source_name": "sample.docx", "raw_text": "题目",
        "curriculum_version": "人教B版", "curriculum_chapter": "函数",
        "curriculum_section": "函数性质", "classification_source": "standard_19",
        "classification_confidence": 1.0, "review_reason": "", "asset_manifest": [{"src": "a.png"}],
    })
    assert draft.curriculum_chapter == "函数"
    assert draft.classification_source == "standard_19"
    assert draft.classification_confidence == 1.0
    assert draft.asset_manifest == [{"src": "a.png"}]
