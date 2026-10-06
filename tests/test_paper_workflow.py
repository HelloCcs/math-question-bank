from question_bank.parser import QuestionDraft
from question_bank.storage import create_export_record, create_import_job, create_paper, get_paper, list_export_records, list_papers, list_questions, save_questions

def _seed(app,count=3):
    db=app.config["DB_CONFIG"]
    import_id=create_import_job(db,source_filename="w.docx",stored_filename="w.docx",preview_relpath="",stats={})
    save_questions(db,import_id,[QuestionDraft(str(i),f"题{i}",[],"","","解答题","w.docx",f"{i}.题") for i in range(1,count+1)])
    return [q["id"] for q in list_questions(db)]

def test_selection_persists_deduplicates_and_saves(client,app):
    ids=_seed(app)
    response=client.post("/selection/add",data={"question_ids":[str(ids[0]),str(ids[1]),str(ids[0])]},follow_redirects=True)
    assert response.status_code==200
    page=client.get("/papers/current"); assert "已选择 2 道题".encode() in page.data
    response=client.post("/papers/current",data={"title":"测试卷","exam_time":"120 分钟","total_score":"150"},follow_redirects=True)
    assert response.status_code==200 and "测试卷".encode() in response.data
    papers=list_papers(app.config["DB_CONFIG"]); assert len(papers)==1
    assert get_paper(app.config["DB_CONFIG"],papers[0]["id"])["question_count"]==2

def test_selection_add_can_return_to_paper_workspace(client, app):
    ids = _seed(app)
    response = client.post(
        "/selection/add",
        data={"question_id": str(ids[0]), "next": "/paper"},
        follow_redirects=False,
    )
    assert response.status_code in {302, 303}
    assert response.headers["Location"].endswith("/paper")


def test_selection_add_ajax_stays_on_current_page(client, app):
    ids = _seed(app)
    response = client.post(
        "/selection/add",
        data={"question_id": str(ids[0]), "next": "/paper"},
        headers={"X-Requested-With": "fetch"},
    )
    assert response.status_code == 200
    assert response.is_json
    assert response.json["ok"] is True
    assert response.json["selected_count"] == 1
    assert len(response.json["selected_questions"]) == 1
    assert response.json["selected_questions"][0]["id"] == ids[0]

def test_paper_requires_metadata_and_can_reorder(client,app):
    ids=_seed(app); client.post("/selection/add",data={"question_ids":ids})
    response=client.post("/papers/current",data={"title":"","exam_time":"","total_score":""},follow_redirects=True)
    assert "必填项".encode() in response.data
    client.post("/papers/current",data={"title":"卷","exam_time":"90","total_score":"100"})
    paper_id=list_papers(app.config["DB_CONFIG"])[0]["id"]
    response=client.post(f"/papers/{paper_id}",data={"title":"卷2","question_ids":f"{ids[2]},{ids[0]}"},follow_redirects=True)
    assert response.status_code==200
    assert [q["id"] for q in get_paper(app.config["DB_CONFIG"],paper_id)["questions"]]==[ids[2],ids[0]]

def test_paper_history_delete(client,app):
    ids=_seed(app,1); client.post("/selection/add",data={"question_ids":ids}); client.post("/papers/current",data={"title":"卷","exam_time":"90","total_score":"100"})
    paper_id=list_papers(app.config["DB_CONFIG"])[0]["id"]
    assert client.get("/papers/history").status_code==200
    client.post(f"/papers/{paper_id}/delete")
    assert list_papers(app.config["DB_CONFIG"])==[]

def test_paper_history_hides_question_mark_garble(client, app):
    ids = _seed(app, 1)
    paper_id = create_paper(
        app.config["DB_CONFIG"],
        {"title": "????????", "exam_time": "120 ??", "total_score": "150 ??"},
        ids,
    )
    create_export_record(app.config["DB_CONFIG"], paper_id, "Word ZIP", "paper.zip", 1)

    page = client.get("/papers/history")
    assert page.status_code == 200
    assert b"????????" not in page.data
    assert b"120 ??" not in page.data
    assert "\u672a\u547d\u540d\u8bd5\u5377".encode() in page.data
    assert "120 \u5206\u949f".encode() in page.data


def test_smart_paper_workspace_shows_library_and_selected_preview(client, app):
    ids = _seed(app, 3)
    page = client.get("/paper")
    assert page.status_code == 200
    assert "\u9898\u5e93\u9898\u76ee".encode() in page.data
    assert b"paper-empty" in page.data
    client.post("/selection/add", data={"question_id": str(ids[0])})
    page = client.get("/paper")
    assert "\u8bd5\u5377\u9884\u89c8".encode() in page.data
    assert "\u5df2\u52a0\u5165\u7ec4\u5377".encode() in page.data
    assert "\u7b2c1\u9898".encode() in page.data

def test_smart_paper_workspace_can_remove_and_move(client, app):
    ids = _seed(app, 3)
    client.post("/selection/add", data={"question_ids": [str(ids[0]), str(ids[1])]})
    client.post("/selection/move", data={"question_id": str(ids[1]), "direction": "up"}, follow_redirects=True)
    client.post("/selection/remove", data={"question_id": str(ids[0])}, follow_redirects=True)
    page = client.get("/paper")
    assert "\u5f53\u524d 1 \u9053\u9898".encode() in page.data


def test_export_history_records_and_reexports(client, app, monkeypatch):
    ids = _seed(app, 1)
    client.post("/selection/add", data={"question_ids": ids})
    client.post("/papers/current", data={"title": "圆锥曲线综合卷", "exam_time": "90 分钟", "total_score": "100"})
    paper_id = list_papers(app.config["DB_CONFIG"])[0]["id"]

    def fake_export(_model, path):
        path.write_bytes(b"zip")
        return path

    monkeypatch.setattr("question_bank.app.export_paper_word_zip", fake_export)
    response = client.get(f"/papers/{paper_id}/export/word")
    assert response.status_code == 200
    records = list_export_records(app.config["DB_CONFIG"])
    assert len(records) == 1
    assert records[0]["title"] == "圆锥曲线综合卷"
    assert records[0]["export_format"] == "Word ZIP"
    history = client.get("/papers/history")
    assert "\u5df2\u5bfc\u51fa\u8bd5\u5377".encode() in history.data
    assert "\u518d\u6b21\u5bfc\u51fa Word ZIP".encode() in history.data
