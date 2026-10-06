from __future__ import annotations

import json
import shutil
import uuid
from dataclasses import asdict, replace
from datetime import datetime
from html import escape
from pathlib import Path
from urllib.parse import urlparse

from flask import (
    Flask,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    session,
    send_file,
    send_from_directory,
    url_for,
)

from .exporter import (
    export_questions_to_docx,
    export_questions_to_image_zip,
    export_questions_to_pdf,
    export_questions_to_xlsx,
    export_paper_word_zip,
    export_paper_pdf,
)
from .ocr import create_ocr_preview, get_ocr_status, run_ocr_file
from .parser import (
    QuestionDraft,
    analyze_docx_assets,
    parse_docx_questions,
    parse_docx_single_question,
    parse_plain_text_questions,
    parse_preview_questions,
    parse_preview_single_question,
)
from .storage import (
    create_import_job,
    dashboard_stats,
    delete_question,
    get_questions_by_ids,
    init_db,
    list_imports,
    list_questions,
    load_database_config,
    save_questions,
    storage_label,
    update_question,
    clear_library_data,
    create_paper,
    update_paper,
    get_paper,
    list_papers,
    delete_paper,
    create_export_record,
    list_export_records,
)
from .taxonomy import DIFFICULTIES, MAJOR_CATEGORIES, QUESTION_TYPES, enrich_question_draft, enrich_questions
from .word_preview import convert_docx_to_preview
from .assets import build_asset_manifest
from .ai_classifier import classify_import_questions, load_ai_config, masked_ai_status, save_ai_config
from .papers import build_export_model


PAPER_QUESTION_TYPES = ["单选题", "多选题", "填空题", "解答题"]


def create_app(config: dict | None = None) -> Flask:
    package_dir = Path(__file__).resolve().parent
    app = Flask(
        __name__,
        template_folder=str(package_dir / "templates"),
        static_folder=str(package_dir / "static"),
    )
    project_root = Path.cwd()
    data_root = project_root / "data"
    app.config.update(
        SECRET_KEY="local-question-bank",
        PROJECT_ROOT=project_root,
        DATA_ROOT=data_root,
        DB_CONFIG=load_database_config(project_root),
        ORIGINAL_DIR=data_root / "originals",
        PREVIEW_DIR=data_root / "previews",
        EXPORT_DIR=data_root / "exports",
        PENDING_DIR=data_root / "pending",
        MAX_CONTENT_LENGTH=100 * 1024 * 1024,
    )
    if config:
        app.config.update(config)
        if "DATA_ROOT" in config and "PENDING_DIR" not in config:
            app.config["PENDING_DIR"] = Path(app.config["DATA_ROOT"]) / "pending"

    _ensure_dirs(app)
    try:
        init_db(app.config["DB_CONFIG"])
        app.config["DB_READY"] = True
        app.config["DB_ERROR"] = ""
    except Exception as exc:
        app.config["DB_READY"] = False
        app.config["DB_ERROR"] = str(exc)

    @app.get("/")
    def index():
        if not app.config["DB_READY"]:
            return _render_db_setup(app)
        imports = list_imports(app.config["DB_CONFIG"])
        return render_template(
            "index.html",
            stats=dashboard_stats(app.config["DB_CONFIG"]),
            imports=imports,
            grouped_imports=_group_imports_by_month(imports),
            questions=list_questions(app.config["DB_CONFIG"], limit=12),
            storage=storage_label(app.config["DB_CONFIG"]),
        )

    @app.route("/settings/ai", methods=["GET", "POST"])
    def ai_settings():
        ai_config = load_ai_config(app.config["PROJECT_ROOT"])
        if request.method == "POST":
            save_ai_config(
                app.config["PROJECT_ROOT"],
                {
                    "enabled": request.form.get("enabled", "") == "1",
                    "base_url": request.form.get("base_url", "").strip(),
                    "api_key": request.form.get("api_key", "").strip(),
                    "model": request.form.get("model", "").strip(),
                    "timeout": request.form.get("timeout", "20").strip(),
                },
            )
            flash("AI 配置已保存。", "success")
            return redirect(url_for("ai_settings"))
        return render_template(
            "ai_settings.html",
            ai_config=ai_config,
            ai_status=masked_ai_status(ai_config),
        )

    @app.post("/admin/reset")
    def reset_library():
        if not app.config["DB_READY"]:
            return _render_db_setup(app)
        clear_library_data(app.config["DB_CONFIG"])
        for key in ("ORIGINAL_DIR", "PREVIEW_DIR", "EXPORT_DIR", "PENDING_DIR"):
            target = Path(app.config[key])
            if target.exists():
                shutil.rmtree(target)
            target.mkdir(parents=True, exist_ok=True)
        flash("题库数据和导入记录已清空。", "success")
        return redirect(url_for("index"))

    @app.route("/import", methods=["GET", "POST"])
    def import_docx():
        if not app.config["DB_READY"]:
            return _render_db_setup(app)
        if request.method == "GET":
            return render_template(
                "import.html",
                ocr_status=get_ocr_status(),
                ai_status=masked_ai_status(load_ai_config(app.config["PROJECT_ROOT"])),
                categories=MAJOR_CATEGORIES,
                difficulties=DIFFICULTIES,
            )

        upload = request.files.get("file")
        if not upload or not upload.filename:
            flash("请先选择一个文件。", "error")
            return redirect(url_for("import_docx"))
        suffix = Path(upload.filename).suffix.lower()
        if suffix not in {".docx", ".pdf", ".png", ".jpg", ".jpeg"}:
            flash("当前版本支持 .docx、.pdf、.png、.jpg、.jpeg。", "error")
            return redirect(url_for("import_docx"))

        token = datetime.now().strftime("%Y%m%d%H%M%S") + "-" + uuid.uuid4().hex[:8]
        stored_filename = f"{token}{suffix}"
        original_path = app.config["ORIGINAL_DIR"] / stored_filename
        upload.save(original_path)

        if suffix == ".docx":
            return _import_docx_file(app, upload.filename, original_path, token, stored_filename)
        return _import_ocr_file(app, upload.filename, original_path, token, stored_filename)

    @app.route("/import/confirm/<token>", methods=["GET", "POST"])
    def confirm_import(token: str):
        if not app.config["DB_READY"]:
            return _render_db_setup(app)
        pending = _load_pending_import(app, token)
        if pending is None:
            flash("这次导入记录已失效，请重新选择文件。", "error")
            return redirect(url_for("import_docx"))

        if request.method == "GET":
            return render_template(
                "import_confirm.html",
                token=token,
                pending=pending,
                categories=MAJOR_CATEGORIES,
                difficulties=DIFFICULTIES,
                question_types=QUESTION_TYPES,
                category_counts=_pending_category_counts(pending["questions"]),
                ai_import_status=_pending_ai_status(pending),
            )

        if request.form.get("action") == "cancel":
            _delete_pending_import(app, token)
            flash("已取消导入，本次识别结果没有写入题库。", "warning")
            return redirect(url_for("import_docx"))

        questions = _questions_from_confirmation_form(pending["questions"], request.form)
        if not questions:
            flash("请至少保留一道题再确认导入。", "error")
            return redirect(url_for("confirm_import", token=token))

        stats = dict(pending["stats"])
        stats["question_count"] = len(questions)
        import_id = _save_import(
            app,
            pending["source_filename"],
            pending["stored_filename"],
            pending["preview_relpath"],
            stats,
            questions,
        )
        _delete_pending_import(app, token)
        flash(f"确认导入完成：已写入 {len(questions)} 道题。", "success")
        return redirect(url_for("questions", import_id=import_id))

    @app.route("/manual", methods=["GET", "POST"])
    def manual_import():
        if not app.config["DB_READY"]:
            return _render_db_setup(app)
        if request.method == "GET":
            return render_template(
                "manual.html",
                categories=MAJOR_CATEGORIES,
                difficulties=DIFFICULTIES,
                question_types=QUESTION_TYPES,
            )

        upload = request.files.get("file")
        if upload and upload.filename:
            suffix = Path(upload.filename).suffix.lower()
            if suffix not in {".docx", ".pdf", ".png", ".jpg", ".jpeg"}:
                flash("单题导入支持 .docx、.pdf、.png、.jpg、.jpeg。", "error")
                return redirect(url_for("manual_import"))
            token = datetime.now().strftime("%Y%m%d%H%M%S") + "-" + uuid.uuid4().hex[:8]
            stored_filename = f"manual-{token}{suffix}"
            original_path = app.config["ORIGINAL_DIR"] / stored_filename
            upload.save(original_path)
            if suffix == ".docx":
                return _import_docx_file(
                    app,
                    upload.filename,
                    original_path,
                    token,
                    stored_filename,
                    single_question_mode=True,
                )
            return _import_ocr_file(app, upload.filename, original_path, token, stored_filename)

        stem = request.form.get("stem", "").strip()
        stem_image = request.files.get("stem_image")
        if not stem and not (stem_image and stem_image.filename):
            flash("题干或题目图片至少填写一项。", "error")
            return redirect(url_for("manual_import"))

        token = datetime.now().strftime("%Y%m%d%H%M%S") + "-" + uuid.uuid4().hex[:8]
        preview_dir = Path(app.config["PREVIEW_DIR"]) / f"manual-{token}"
        preview_dir.mkdir(parents=True, exist_ok=True)
        stem_image_relpath = _save_manual_image(stem_image, preview_dir, "stem")
        question = _build_manual_question(request.form, stem_image_relpath=stem_image_relpath)
        preview_relpath = _write_manual_preview(app, token, question)
        import_id = _save_import(
            app,
            "单题导入",
            f"manual-{token}.html",
            preview_relpath,
            {"import_mode": "manual", "question_count": 1},
            [question],
        )
        flash("单题导入完成。", "success")
        return redirect(url_for("questions", import_id=import_id))

    @app.route("/paper", methods=["GET", "POST"])
    def paper_builder():
        if not app.config["DB_READY"]:
            return _render_db_setup(app)

        selected_questions: list[dict] = get_questions_by_ids(app.config["DB_CONFIG"], _session_question_ids())
        shortages: list[str] = []
        export_ids = ""
        form_state = _paper_form_state(request.form if request.method == "POST" else {})
        if request.method == "POST" and request.form.get("mode") == "random":
            total_requested = sum(form_state["counts"].values())
            if total_requested <= 0:
                flash("请至少设置一种题型的数量。", "error")
                return redirect(url_for("paper_builder"))
            if total_requested > 100:
                flash("智能组卷最多 100 道题。", "error")
                return redirect(url_for("paper_builder"))
            selected_questions, shortages = _select_paper_questions(app, form_state)
            if not selected_questions:
                flash("没有匹配到可组卷的题目，请调整条件。", "error")
                return redirect(url_for("paper_builder"))
            export_ids = ",".join(str(question["id"]) for question in selected_questions)

        return render_template(
            "paper.html",
            library_questions=list_questions(
                app.config["DB_CONFIG"],
                keyword=request.args.get("q", "").strip(),
                question_type=request.args.get("type", "").strip(),
                category_major=request.args.get("category", "").strip(),
                difficulty=request.args.get("difficulty", "").strip(),
                limit=300,
            ),
            categories=MAJOR_CATEGORIES,
            difficulties=DIFFICULTIES,
            question_types=PAPER_QUESTION_TYPES,
            form_state=form_state,
            selected_questions=selected_questions,
            shortages=shortages,
            export_ids=export_ids,
            selected_ids=set(_session_question_ids()),
            workspace_keyword=request.args.get("q", "").strip(),
            workspace_type=request.args.get("type", "").strip(),
            workspace_category=request.args.get("category", "").strip(),
            workspace_difficulty=request.args.get("difficulty", "").strip(),
        )

    @app.get("/questions")
    def questions():
        if not app.config["DB_READY"]:
            return _render_db_setup(app)
        keyword = request.args.get("q", "").strip()
        question_type = request.args.get("type", "").strip()
        category_major = request.args.get("category", "").strip()
        category_minor = request.args.get("minor", "").strip()
        difficulty = request.args.get("difficulty", "").strip()
        import_id_raw = request.args.get("import_id", "").strip()
        import_id = int(import_id_raw) if import_id_raw.isdigit() else None
        rows = list_questions(
            app.config["DB_CONFIG"],
            keyword=keyword,
            question_type=question_type,
            category_major=category_major,
            category_minor=category_minor,
            difficulty=difficulty,
            import_id=import_id,
            limit=1000,
        )
        stats = dashboard_stats(app.config["DB_CONFIG"])
        return render_template(
            "questions.html",
            questions=rows,
            imports=list_imports(app.config["DB_CONFIG"]),
            type_counts=stats["type_counts"],
            category_counts=stats["category_counts"],
            difficulty_counts=stats["difficulty_counts"],
            category_total=sum(item["count"] for item in stats["category_counts"]),
            difficulty_total=sum(item["count"] for item in stats["difficulty_counts"]),
            categories=MAJOR_CATEGORIES,
            difficulties=DIFFICULTIES,
            question_types=QUESTION_TYPES,
            keyword=keyword,
            selected_type=question_type,
            selected_category=category_major,
            selected_minor=category_minor,
            selected_difficulty=difficulty,
            selected_import=import_id_raw,
            selected_ids=set(_session_question_ids()),
        )

    @app.post("/selection/add")
    def selection_add():
        existing = _session_question_ids()
        incoming = _parse_ids(",".join(request.form.getlist("question_ids")))
        if request.form.get("question_id", "").isdigit(): incoming.append(int(request.form["question_id"]))
        valid = [row["id"] for row in get_questions_by_ids(app.config["DB_CONFIG"], incoming)]
        merged = list(dict.fromkeys(existing + valid))
        is_ajax = request.headers.get("X-Requested-With") == "fetch"
        if len(merged) > 100:
            message = "\u6bcf\u4efd\u7ec4\u5377\u6700\u591a\u5305\u542b 100 \u9053\u9898\u3002"
            if is_ajax:
                return jsonify({"ok": False, "message": message, "selected_count": len(existing)}), 400
            flash(message, "error")
        else:
            session["paper_selection"] = merged
            message = f"\u5df2\u52a0\u5165\u7ec4\u5377\uff0c\u5f53\u524d\u5171 {len(merged)} \u9053\u9898\u3002"
            if is_ajax:
                return jsonify({
                    "ok": True,
                    "message": message,
                    "selected_count": len(merged),
                    "added_ids": valid,
                    "selected_questions": [_paper_selection_payload(row) for row in get_questions_by_ids(app.config["DB_CONFIG"], merged)],
                })
            flash(message, "success")
        return redirect(_safe_redirect_target(request.form.get("next"), request.referrer, default=url_for("questions")))

    @app.post("/selection/clear")
    def selection_clear():
        session.pop("paper_selection", None); flash("当前组卷已清空。", "success")
        return redirect(request.referrer or url_for("questions"))

    @app.post("/selection/remove")
    def selection_remove():
        remove_id = request.form.get("question_id", "")
        ids = [value for value in _session_question_ids() if str(value) != remove_id]
        session["paper_selection"] = ids
        if request.headers.get("X-Requested-With") == "fetch":
            return jsonify({
                "ok": True,
                "selected_count": len(ids),
                "selected_questions": [_paper_selection_payload(row) for row in get_questions_by_ids(app.config["DB_CONFIG"], ids)],
            })
        return redirect(request.referrer or url_for("paper_builder"))

    @app.post("/selection/move")
    def selection_move():
        ids = _session_question_ids()
        raw_id = request.form.get("question_id", "")
        direction = request.form.get("direction", "")
        try:
            index = ids.index(int(raw_id))
        except (ValueError, TypeError):
            return redirect(request.referrer or url_for("paper_builder"))
        target = index - 1 if direction == "up" else index + 1
        if 0 <= target < len(ids):
            ids[index], ids[target] = ids[target], ids[index]
            session["paper_selection"] = ids
        if request.headers.get("X-Requested-With") == "fetch":
            return jsonify({
                "ok": True,
                "selected_count": len(ids),
                "selected_questions": [_paper_selection_payload(row) for row in get_questions_by_ids(app.config["DB_CONFIG"], ids)],
            })
        return redirect(request.referrer or url_for("paper_builder"))

    @app.route("/papers/current", methods=["GET", "POST"])
    def current_paper():
        ids = _session_question_ids()
        if request.method == "POST":
            if not ids:
                flash("请先从题库选择题目。", "error"); return redirect(url_for("questions"))
            title=request.form.get("title","").strip(); exam_time=request.form.get("exam_time","").strip(); total_score=request.form.get("total_score","").strip()
            if not title or not exam_time or not total_score:
                flash("试卷标题、考试时间和满分为必填项。", "error")
            else:
                paper_id=create_paper(app.config["DB_CONFIG"], {k:request.form.get(k,"") for k in ("title","exam_time","total_score","school","grade","class_name","instructions")}|{"source":"manual"}, ids)
                session.pop("paper_selection",None); flash("组卷已保存。", "success"); return redirect(url_for("paper_detail",paper_id=paper_id))
        return render_template("paper_current.html",questions=get_questions_by_ids(app.config["DB_CONFIG"],ids),selected_count=len(ids))

    @app.get("/papers/history")
    def paper_history():
        return render_template(
            "paper_history.html",
            papers=list_papers(app.config["DB_CONFIG"]),
            exports=list_export_records(app.config["DB_CONFIG"]),
        )

    @app.route("/papers/<int:paper_id>", methods=["GET", "POST"])
    def paper_detail(paper_id):
        paper=get_paper(app.config["DB_CONFIG"],paper_id)
        if paper is None: flash("组卷记录不存在。","error"); return redirect(url_for("paper_history"))
        if request.method == "POST":
            ids=_parse_ids(request.form.get("question_ids",""))
            update_paper(app.config["DB_CONFIG"],paper_id,{k:request.form.get(k,paper.get(k,"")) for k in ("title","exam_time","total_score","school","grade","class_name","instructions")},ids)
            flash("组卷已更新。","success"); return redirect(url_for("paper_detail",paper_id=paper_id))
        return render_template("paper_detail.html",paper=paper)

    @app.post("/papers/<int:paper_id>/delete")
    def paper_delete(paper_id):
        delete_paper(app.config["DB_CONFIG"],paper_id); flash("组卷历史已删除。","success"); return redirect(url_for("paper_history"))

    @app.get("/papers/<int:paper_id>/export/<fmt>")
    def paper_export(paper_id,fmt):
        paper=get_paper(app.config["DB_CONFIG"],paper_id)
        if paper is None: flash("组卷记录不存在。","error"); return redirect(url_for("paper_history"))
        try: model=build_export_model(paper)
        except ValueError as exc: flash(str(exc),"error"); return redirect(url_for("paper_detail",paper_id=paper_id))
        stamp=datetime.now().strftime("%Y%m%d-%H%M%S"); safe="".join(ch if ch.isalnum() or ch in "-_" else "-" for ch in model.title).strip("-") or "试卷"
        if fmt=="word": exporter,name,mime=export_paper_word_zip,f"{safe}-{stamp}.zip","application/zip"
        elif fmt=="pdf": exporter,name,mime=export_paper_pdf,f"{safe}-{stamp}.pdf","application/pdf"
        else: flash("不支持的组卷导出格式。","error"); return redirect(url_for("paper_detail",paper_id=paper_id))
        path=app.config["EXPORT_DIR"]/name
        try: exporter(model,path)
        except Exception as exc: flash(f"导出失败：{exc}","error"); return redirect(url_for("paper_detail",paper_id=paper_id))
        create_export_record(app.config["DB_CONFIG"], paper_id, "Word ZIP" if fmt == "word" else "PDF", name, len(model.questions))
        return send_file(path,as_attachment=True,download_name=name,mimetype=mime)

    @app.post("/questions/<int:question_id>/update")
    def update_question_route(question_id: int):
        if not app.config["DB_READY"]:
            return _render_db_setup(app)
        data = {
            field: request.form.get(field, "").strip()
            for field in ("number", "question_type", "category_major", "category_minor", "difficulty")
            if field in request.form
        }
        for field in ("stem", "options", "answer", "analysis"):
            if field not in request.form:
                continue
            value = request.form.get(field, "").strip()
            original_value = request.form.get(f"original_{field}", "").strip()
            if value == original_value:
                continue
            data[field] = value
            if field == "options":
                data["options_html"] = _text_to_paragraph_html("\n".join(_split_option_lines(value)))
            else:
                data[f"{field}_html"] = _text_to_paragraph_html(value)
        update_question(app.config["DB_CONFIG"], question_id, data)
        flash("题目已保存。", "success")
        return redirect(request.referrer or url_for("questions"))

    @app.post("/questions/<int:question_id>/delete")
    def delete_question_route(question_id: int):
        if not app.config["DB_READY"]:
            return _render_db_setup(app)
        delete_question(app.config["DB_CONFIG"], question_id)
        flash("题目已删除。", "success")
        return redirect(request.referrer or url_for("questions"))

    @app.get("/export/<fmt>")
    def export(fmt: str):
        if not app.config["DB_READY"]:
            return _render_db_setup(app)
        ids = _parse_ids(request.args.get("ids", ""))
        if ids:
            questions_to_export = get_questions_by_ids(app.config["DB_CONFIG"], ids)
        else:
            import_id_raw = request.args.get("import_id", "").strip()
            import_id = int(import_id_raw) if import_id_raw.isdigit() else None
            questions_to_export = list_questions(
                app.config["DB_CONFIG"],
                keyword=request.args.get("q", "").strip(),
                question_type=request.args.get("type", "").strip(),
                category_major=request.args.get("category", "").strip(),
                category_minor=request.args.get("minor", "").strip(),
                difficulty=request.args.get("difficulty", "").strip(),
                import_id=import_id,
                limit=10000,
            )
        if not questions_to_export:
            flash("没有可导出的题目。", "error")
            return redirect(url_for("questions"))

        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        exporters = {
            "xlsx": (export_questions_to_xlsx, f"题库导出-{timestamp}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            "docx": (export_questions_to_docx, f"题库导出-{timestamp}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
            "pdf": (export_questions_to_pdf, f"题库导出-{timestamp}.pdf", "application/pdf"),
            "stems-docx": (lambda rows, path: export_questions_to_docx(rows, path, stems_only=True), f"题库题干-{timestamp}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
            "stems-pdf": (lambda rows, path: export_questions_to_pdf(rows, path, stems_only=True), f"题库题干-{timestamp}.pdf", "application/pdf"),
            "images": (export_questions_to_image_zip, f"题目图片-{timestamp}.zip", "application/zip"),
        }
        if fmt not in exporters:
            flash("不支持的导出格式。", "error")
            return redirect(url_for("questions"))

        exporter, filename, mimetype = exporters[fmt]
        output_path = app.config["EXPORT_DIR"] / filename
        try:
            exporter(questions_to_export, output_path)
        except Exception as exc:
            flash(f"导出失败：{exc}", "error")
            return redirect(url_for("questions"))
        return send_file(output_path, as_attachment=True, download_name=filename, mimetype=mimetype)

    @app.get("/data/<path:filename>")
    def data_file(filename: str):
        path = Path(app.config["DATA_ROOT"]) / filename
        if path.suffix.lower() in {".htm", ".html"} and path.exists():
            return app.response_class(
                path.read_text(encoding="utf-8", errors="ignore"),
                content_type="text/html; charset=utf-8",
            )
        return send_from_directory(app.config["DATA_ROOT"], filename)

    @app.template_filter("stats")
    def stats_filter(value: str | dict) -> dict:
        if isinstance(value, dict):
            return value
        try:
            return json.loads(value or "{}")
        except json.JSONDecodeError:
            return {}

    return app


def _ensure_dirs(app: Flask) -> None:
    for key in ("DATA_ROOT", "ORIGINAL_DIR", "PREVIEW_DIR", "EXPORT_DIR", "PENDING_DIR"):
        Path(app.config[key]).mkdir(parents=True, exist_ok=True)


def _parse_ids(raw: str) -> list[int]:
    ids: list[int] = []
    for item in raw.split(","):
        item = item.strip()
        if item.isdigit():
            ids.append(int(item))
    return ids


def _safe_redirect_target(*candidates: str | None, default: str) -> str:
    for candidate in candidates:
        if not candidate:
            continue
        parsed = urlparse(candidate)
        if not parsed.scheme and not parsed.netloc and candidate.startswith("/"):
            return candidate
        if parsed.scheme in {"http", "https"} and parsed.netloc == request.host:
            return parsed.path + (f"?{parsed.query}" if parsed.query else "")
    return default


def _session_question_ids() -> list[int]:
    values = session.get("paper_selection", [])
    return [int(value) for value in values if str(value).isdigit()][:100]


def _import_docx_file(
    app: Flask,
    source_filename: str,
    original_path: Path,
    token: str,
    stored_filename: str,
    *,
    single_question_mode: bool = False,
):
    preview_relpath = f"previews/{token}/preview.htm"
    preview_dir = app.config["PREVIEW_DIR"] / token
    summary = analyze_docx_assets(original_path)
    plain_questions = (
        parse_docx_single_question(original_path, source_name=source_filename)
        if single_question_mode
        else parse_docx_questions(original_path, source_name=source_filename)
    )

    preview_method = "none"
    preview_path = None
    preview_converter = app.config.get("PREVIEW_CONVERTER", convert_docx_to_preview)
    try:
        preview_path, preview_method = preview_converter(original_path, preview_dir)
    except Exception as exc:
        flash(f"Word 预览转换失败，但题目文本已尝试导入：{exc}", "warning")

    if preview_path and preview_path.exists():
        rich_questions = (
            parse_preview_single_question(
                preview_path,
                source_name=source_filename,
                data_url_prefix=f"/data/previews/{token}",
            )
            if single_question_mode
            else parse_preview_questions(
                preview_path,
                source_name=source_filename,
                data_url_prefix=f"/data/previews/{token}",
            )
        )
        questions = _merge_docx_preview_questions(plain_questions, rich_questions)
    else:
        questions = plain_questions

    stats = {
        "import_mode": "word",
        "media_counts": summary.media_counts,
        "formula_count": summary.formula_count,
        "page_count": summary.page_count,
        "question_count": len(questions),
        "preview_method": preview_method,
    }
    _save_pending_import(app, token, source_filename, stored_filename, preview_relpath, stats, questions)
    flash(f"识别完成：识别到 {len(questions)} 道题，请确认分类后再入库。", "success")
    return redirect(url_for("confirm_import", token=token))


def _merge_docx_preview_questions(
    plain_questions: list[QuestionDraft],
    rich_questions: list[QuestionDraft],
) -> list[QuestionDraft]:
    if not plain_questions:
        return rich_questions
    if not rich_questions:
        return plain_questions

    plain_by_number = {question.number: question for question in plain_questions}
    rich_by_number = {question.number: question for question in rich_questions}
    ordered_numbers = sorted(
        {question.number for question in plain_questions + rich_questions},
        key=lambda value: (_safe_int(value) is None, _safe_int(value) or 0, value),
    )

    merged: list[QuestionDraft] = []
    for number in ordered_numbers:
        plain = plain_by_number.get(number)
        rich = rich_by_number.get(number)
        if plain is None and rich is not None:
            merged.append(rich)
            continue
        if plain is not None and rich is None:
            merged.append(_with_plain_html(plain))
            continue
        if plain is None or rich is None:
            continue
        merged.append(
            replace(
                plain,
                stem_html=rich.stem_html or _paragraph_html(plain.stem),
                options_html=rich.options_html or _options_html(plain.options),
                answer_html=rich.answer_html or _paragraph_html(plain.answer),
                analysis_html=rich.analysis_html or _paragraph_html(plain.analysis),
                raw_html=rich.raw_html or _paragraph_html(plain.raw_text),
            )
        )
    return merged


def _with_plain_html(question: QuestionDraft) -> QuestionDraft:
    return replace(
        question,
        stem_html=_paragraph_html(question.stem),
        options_html=_options_html(question.options),
        answer_html=_paragraph_html(question.answer),
        analysis_html=_paragraph_html(question.analysis),
        raw_html=_paragraph_html(question.raw_text),
    )


def _paragraph_html(text: str) -> str:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    return "".join(f"<p>{escape(line)}</p>" for line in lines)


def _options_html(options: list[str]) -> str:
    return "".join(f"<p>{escape(option)}</p>" for option in options if option)


def _import_ocr_file(app: Flask, source_filename: str, original_path: Path, token: str, stored_filename: str):
    preview_dir = app.config["PREVIEW_DIR"] / token
    ocr_processor = app.config.get("OCR_PROCESSOR", run_ocr_file)
    try:
        result = ocr_processor(original_path, preview_dir)
    except Exception as exc:
        flash(f"OCR 识别失败：{exc}", "error")
        return redirect(url_for("import_docx"))

    preview_path = create_ocr_preview(result, preview_dir)
    questions = parse_plain_text_questions(result.text, source_name=source_filename)
    preview_relpath = f"previews/{token}/{preview_path.name}"
    stats = {
        "import_mode": "ocr",
        "ocr_engine": result.engine,
        "ocr_language": result.language,
        "ocr_warnings": result.warnings,
        "page_count": len(result.pages),
        "question_count": len(questions),
    }
    _save_pending_import(app, token, source_filename, stored_filename, preview_relpath, stats, questions)
    for warning in result.warnings:
        flash(warning, "warning")
    flash(f"OCR 识别完成：识别到 {len(questions)} 道题，请确认分类后再入库。", "success")
    return redirect(url_for("confirm_import", token=token))


def _save_import(
    app: Flask,
    source_filename: str,
    stored_filename: str,
    preview_relpath: str,
    stats: dict,
    questions,
) -> int:
    import_id = create_import_job(
        app.config["DB_CONFIG"],
        source_filename=source_filename,
        stored_filename=stored_filename,
        preview_relpath=preview_relpath,
        stats=stats,
    )
    save_questions(app.config["DB_CONFIG"], import_id, questions)
    return import_id


def _save_pending_import(
    app: Flask,
    token: str,
    source_filename: str,
    stored_filename: str,
    preview_relpath: str,
    stats: dict,
    questions,
) -> None:
    enriched = enrich_questions(list(questions))
    ai_config = load_ai_config(app.config["PROJECT_ROOT"])
    enriched, ai_warnings = classify_import_questions(enriched, ai_config)
    if ai_warnings:
        stats["ai_warnings"] = ai_warnings
        for warning in ai_warnings:
            flash(warning, "warning")
    enriched = [replace(q, asset_manifest=build_asset_manifest(q.raw_html)) for q in enriched]
    payload = {
        "token": token,
        "source_filename": source_filename,
        "stored_filename": stored_filename,
        "preview_relpath": preview_relpath,
        "stats": stats,
        "questions": [asdict(question) for question in enriched],
    }
    _pending_path(app, token).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _load_pending_import(app: Flask, token: str) -> dict | None:
    path = _pending_path(app, token)
    if not path.exists():
        return None
    payload = json.loads(path.read_text(encoding="utf-8"))
    raw_questions = payload.get("questions") or []
    if raw_questions and all(item.get("classification_source") in {"rule", "standard_19", "legacy", ""} for item in raw_questions):
        refreshed = enrich_questions([_draft_from_dict(item) for item in raw_questions])
        payload["questions"] = [asdict(question) for question in refreshed]
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return payload


def _delete_pending_import(app: Flask, token: str) -> None:
    path = _pending_path(app, token)
    if path.exists():
        path.unlink()


def _pending_path(app: Flask, token: str) -> Path:
    safe_token = "".join(ch for ch in token if ch.isalnum() or ch in "-_")
    return Path(app.config["PENDING_DIR"]) / f"{safe_token}.json"


def _pending_category_counts(questions: list[dict]) -> list[dict]:
    counts: dict[str, int] = {}
    for question in questions:
        category = (question.get("category_major") or "未分类").strip() or "未分类"
        counts[category] = counts.get(category, 0) + 1
    return [{"category_major": category, "count": count} for category, count in sorted(counts.items())]


def _pending_ai_status(pending: dict) -> dict:
    questions = pending.get("questions") or []
    stats = pending.get("stats") or {}
    warnings = [str(item) for item in stats.get("ai_warnings") or [] if str(item).strip()]
    total = len(questions)
    ai_count = sum(1 for item in questions if item.get("classification_source") == "ai")
    traced = [
        item for item in questions
        if item.get("classification_source") == "ai" and str(item.get("rule_category_major") or "").strip()
    ]
    changed_count = sum(
        1 for item in traced
        if str(item.get("rule_category_major") or "").strip() != str(item.get("ai_category_major") or item.get("category_major") or "").strip()
    )
    confirmed_count = len(traced) - changed_count

    if warnings:
        return {
            "state": "warning",
            "message": "AI 复识别未完成，当前结果主要来自本地规则。",
            "detail": "；".join(warnings),
            "ai_count": ai_count,
            "total": total,
        }
    if ai_count:
        if traced:
            detail = f"AI改动 {changed_count} 题，AI确认本地分类 {confirmed_count} 题。每题标签会显示规则原判和 AI 新判。"
        else:
            detail = "这次记录没有保存 AI 前的本地分类，只能确认 AI 已调用，无法判断 AI 改动了几题；重新导入后会显示逐题对比。"
        return {
            "state": "success",
            "message": f"AI 复识别完成：{ai_count}/{total} 题已由 AI 判断。",
            "detail": detail,
            "ai_count": ai_count,
            "total": total,
            "changed_count": changed_count,
            "confirmed_count": confirmed_count,
        }
    return {
        "state": "idle",
        "message": "未看到 AI 复识别记录，当前结果来自本地规则。",
        "detail": "请在 AI设置 中填写 API Key 后重新导入，即可触发导入后的 AI 复识别。",
        "ai_count": 0,
        "total": total,
    }


def _group_imports_by_month(imports: list[dict]) -> list[dict]:
    grouped: list[dict] = []
    seen: dict[str, dict] = {}
    for item in imports:
        created_at = str(item.get("created_at") or "")
        month_key = created_at[:7] if len(created_at) >= 7 else "未知时间"
        if month_key != "未知时间":
            year, month = month_key.split("-", 1)
            label = f"{year[2:]}年{int(month)}月"
        else:
            label = month_key
        bucket = seen.get(month_key)
        if bucket is None:
            bucket = {"month_key": month_key, "label": label, "items": []}
            seen[month_key] = bucket
            grouped.append(bucket)
        bucket["items"].append(item)
    return grouped


def _build_manual_question(form, *, stem_image_relpath: str = "") -> QuestionDraft:
    options = _split_option_lines(form.get("options", ""))
    stem_html = _text_to_paragraph_html(form.get("stem", ""))
    if stem_image_relpath:
        stem_html += f'<p><img class="inline-math block-figure" src="/data/{stem_image_relpath}" alt="题目图片"></p>'
    draft = QuestionDraft(
        number=form.get("number", "").strip(),
        stem=form.get("stem", "").strip(),
        options=options,
        answer=form.get("answer", "").strip(),
        analysis=form.get("analysis", "").strip(),
        question_type=form.get("question_type", "").strip(),
        source_name="单题导入",
        raw_text="\n".join(
            item
            for item in [
                form.get("stem", "").strip(),
                "\n".join(options),
                form.get("answer", "").strip(),
                form.get("analysis", "").strip(),
            ]
            if item
        ),
        stem_html=stem_html,
        options_html=_text_to_paragraph_html("\n".join(options)),
        answer_html=_text_to_paragraph_html(form.get("answer", "")),
        analysis_html=_text_to_paragraph_html(form.get("analysis", "")),
        category_major=form.get("category_major", "").strip(),
        category_minor=form.get("category_minor", "").strip(),
        difficulty=form.get("difficulty", "").strip(),
    )
    return enrich_question_draft(draft)


def _write_manual_preview(app: Flask, token: str, question: QuestionDraft) -> str:
    preview_dir = Path(app.config["PREVIEW_DIR"]) / f"manual-{token}"
    preview_dir.mkdir(parents=True, exist_ok=True)
    preview_path = preview_dir / "preview.html"
    body = "\n".join(
        part
        for part in [
            f"<section><div><strong>{escape(question.number or '1')}.</strong></div>{question.stem_html}</section>",
            f"<section>{question.options_html}</section>" if question.options_html else "",
            f"<section><p><strong>答案：</strong></p>{question.answer_html}</section>" if question.answer_html else "",
            f"<section><p><strong>解析：</strong></p>{question.analysis_html}</section>" if question.analysis_html else "",
        ]
        if part
    )
    preview_path.write_text(
        f'<!doctype html><html lang="zh-CN"><meta charset="utf-8"><body>{body}</body></html>',
        encoding="utf-8",
    )
    return f"previews/manual-{token}/preview.html"


def _save_manual_image(upload, preview_dir: Path, prefix: str) -> str:
    if upload is None or not getattr(upload, "filename", ""):
        return ""
    suffix = Path(upload.filename).suffix.lower()
    if suffix not in {".png", ".jpg", ".jpeg"}:
        return ""
    stored_name = f"{prefix}{suffix}"
    target = preview_dir / stored_name
    upload.save(target)
    return f"previews/{preview_dir.name}/{stored_name}"


def _split_option_lines(raw: str) -> list[str]:
    return [line.strip() for line in raw.splitlines() if line.strip()]


def _text_to_paragraph_html(raw: str) -> str:
    lines = [line.strip() for line in raw.splitlines() if line.strip()]
    return "".join(f"<p>{escape(line)}</p>" for line in lines)


def _paper_form_state(form) -> dict:
    selected_categories = _form_getlist(form, "categories")
    if not selected_categories:
        legacy_category = form.get("category_major", "").strip() if form else ""
        selected_categories = [legacy_category] if legacy_category else []
    counts = {
        question_type: _safe_int(form.get(f"count_{question_type}", "0"))
        for question_type in PAPER_QUESTION_TYPES
    }
    return {
        "category_major": selected_categories[0] if len(selected_categories) == 1 else "",
        "categories": selected_categories,
        "category_minor": form.get("category_minor", "").strip() if form else "",
        "difficulty": form.get("difficulty", "").strip() if form else "",
        "counts": counts,
    }


def _select_paper_questions(app: Flask, form_state: dict) -> tuple[list[dict], list[str]]:
    selected: list[dict] = []
    shortages: list[str] = []
    used_ids: set[int] = set()
    for question_type in PAPER_QUESTION_TYPES:
        requested = form_state["counts"].get(question_type, 0)
        if requested <= 0:
            continue
        candidates = _paper_candidates(app, question_type, form_state)
        picked = []
        for question in candidates:
            if question["id"] in used_ids:
                continue
            picked.append(question)
            used_ids.add(question["id"])
            if len(picked) >= requested:
                break
        if len(picked) < requested:
            shortages.append(f"{question_type} 需要 {requested} 道，当前匹配 {len(picked)} 道。")
        selected.extend(picked)
    return selected[:100], shortages


def _paper_candidates(app: Flask, question_type: str, form_state: dict) -> list[dict]:
    categories = list(form_state.get("categories") or [])
    if not categories:
        return list_questions(
            app.config["DB_CONFIG"],
            question_type=question_type,
            category_minor=form_state["category_minor"],
            difficulty=form_state["difficulty"],
            limit=10000,
        )

    grouped: list[list[dict]] = [
        list_questions(
            app.config["DB_CONFIG"],
            question_type=question_type,
            category_major=category,
            category_minor=form_state["category_minor"],
            difficulty=form_state["difficulty"],
            limit=10000,
        )
        for category in categories
    ]
    mixed: list[dict] = []
    max_len = max((len(group) for group in grouped), default=0)
    for index in range(max_len):
        for group in grouped:
            if index < len(group):
                mixed.append(group[index])
    return mixed


def _safe_int(raw: object) -> int:
    try:
        value = int(str(raw or "0").strip())
    except ValueError:
        return 0
    return max(value, 0)


def _form_getlist(form, key: str) -> list[str]:
    if not form:
        return []
    if hasattr(form, "getlist"):
        values = form.getlist(key)
    else:
        raw = form.get(key, [])
        values = raw if isinstance(raw, list) else [raw]
    return [str(value).strip() for value in values if str(value).strip()]


def _questions_from_confirmation_form(question_rows: list[dict], form) -> list:
    questions = []
    for index, row in enumerate(question_rows):
        if form.get(f"include_{index}") != "1":
            continue
        data = dict(row)
        for field in ("number", "question_type", "category_major", "category_minor", "difficulty"):
            form_value = form.get(f"{field}_{index}")
            if form_value is not None:
                data[field] = form_value.strip()
        questions.append(_draft_from_dict(data))
    return questions


def _draft_from_dict(data: dict):
    from .parser import QuestionDraft

    return QuestionDraft(
        number=data.get("number", ""),
        stem=data.get("stem", ""),
        options=list(data.get("options") or []),
        answer=data.get("answer", ""),
        analysis=data.get("analysis", ""),
        question_type=data.get("question_type", ""),
        source_name=data.get("source_name", ""),
        raw_text=data.get("raw_text", ""),
        stem_html=data.get("stem_html", ""),
        options_html=data.get("options_html", ""),
        answer_html=data.get("answer_html", ""),
        analysis_html=data.get("analysis_html", ""),
        raw_html=data.get("raw_html", ""),
        category_major=data.get("category_major", ""),
        category_minor=data.get("category_minor", ""),
        difficulty=data.get("difficulty", ""),
        curriculum_version=data.get("curriculum_version", "人教B版"),
        curriculum_chapter=data.get("curriculum_chapter", ""),
        curriculum_section=data.get("curriculum_section", ""),
        classification_source=data.get("classification_source", "rule"),
        classification_confidence=data.get("classification_confidence"),
        review_reason=data.get("review_reason", ""),
        rule_category_major=data.get("rule_category_major", ""),
        rule_category_minor=data.get("rule_category_minor", ""),
        ai_category_major=data.get("ai_category_major", ""),
        ai_category_minor=data.get("ai_category_minor", ""),
        ai_review_note=data.get("ai_review_note", ""),
        asset_manifest=data.get("asset_manifest") or [],
    )


def _paper_selection_payload(question: dict) -> dict:
    return {
        "id": question.get("id"),
        "number": question.get("number") or "",
        "question_type": question.get("question_type") or "",
        "stem": question.get("stem") or "",
        "stem_html": question.get("stem_html") or "",
        "options": question.get("options") or "",
        "options_html": question.get("options_html") or "",
    }


def _render_db_setup(app: Flask):
    return render_template(
        "db_setup.html",
        storage=storage_label(app.config["DB_CONFIG"]),
        error=app.config["DB_ERROR"],
    )
