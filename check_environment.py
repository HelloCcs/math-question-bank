from __future__ import annotations
import importlib.util, socket, sys
from pathlib import Path
from question_bank.ai_classifier import load_ai_config
from question_bank.ocr import get_ocr_status
from question_bank.storage import database_health, init_db, load_database_config

REQUIRED=["flask","docx","openpyxl","PIL","bs4","fitz"]

def run_check(root: Path):
    blocking=[]; optional=[]; ok=[]
    (blocking if sys.version_info<(3,11) else ok).append("Python 需要 3.11 或以上版本" if sys.version_info<(3,11) else f"Python {sys.version.split()[0]}")
    missing=[name for name in REQUIRED if importlib.util.find_spec(name) is None]
    (blocking if missing else ok).append("缺少依赖："+", ".join(missing) if missing else "Python 必需依赖完整")
    try:
        config=load_database_config(root); init_db(config); health=database_health(config)
        (blocking if health["integrity"]!="ok" else ok).append("SQLite 完整性检查失败" if health["integrity"]!="ok" else "SQLite 可读写且完整")
    except Exception as exc: blocking.append(f"SQLite 不可用：{exc}")
    office=Path(r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE")
    (ok if office.exists() else optional).append("Microsoft Word 可用" if office.exists() else "未找到 Microsoft Word，PDF 导出可能不可用")
    ocr=get_ocr_status(); (ok if ocr["available"] else optional).extend(["Tesseract OCR 可用"] if ocr["available"] else ocr["warnings"])
    ai=load_ai_config(root)
    if ai.enabled and not (ai.base_url and ai.api_key and ai.model): optional.append("AI 已启用但配置不完整")
    elif ai.enabled: ok.append("AI 分类已配置")
    else: optional.append("AI 未启用，将使用本地分类")
    try: s=socket.socket(); s.bind(("127.0.0.1",5008)); s.close(); ok.append("端口 5008 可用")
    except OSError: optional.append("端口 5008 已占用；若系统正在运行可忽略")
    lines=["[通过] "+x for x in ok]+["[可选] "+x for x in optional]+["[阻断] "+x for x in blocking]
    return (1 if blocking else 2 if optional else 0),lines

if __name__=="__main__":
    root=Path(__file__).resolve().parent; code,lines=run_check(root); report=root/"environment-report.txt"
    report.write_text("\n".join(lines)+"\n",encoding="utf-8"); print("\n".join(lines)); print(f"报告：{report}"); raise SystemExit(code)
