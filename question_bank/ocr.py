from __future__ import annotations

import html
import os
import shutil
from dataclasses import dataclass
from pathlib import Path

import fitz
import pytesseract
from PIL import Image, ImageOps


COMMON_TESSERACT_PATHS = [
    Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
    Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
    Path.cwd() / "runtime" / "tesseract" / "tesseract.exe",
    Path.cwd() / "tesseract" / "tesseract.exe",
]
PROJECT_TESSDATA = Path.cwd() / "data" / "tessdata"
RUNTIME_TESSDATA = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "QuestionBankOCR" / "tessdata"


@dataclass(frozen=True)
class OcrResult:
    text: str
    pages: list[dict]
    engine: str
    language: str
    warnings: list[str]


def run_ocr_file(path: str | Path, output_dir: str | Path) -> OcrResult:
    source = Path(path)
    target_dir = Path(output_dir)
    target_dir.mkdir(parents=True, exist_ok=True)
    tesseract_cmd = find_tesseract()
    if not tesseract_cmd:
        raise RuntimeError("未找到 tesseract.exe，无法执行 OCR。")
    pytesseract.pytesseract.tesseract_cmd = str(tesseract_cmd)
    tessdata_config = get_tessdata_config()
    languages = list_project_languages() or pytesseract.get_languages(config="")
    language, warnings = choose_tesseract_language(languages)

    if source.suffix.lower() == ".pdf":
        page_images = _render_pdf_pages(source, target_dir)
    else:
        page_images = [_prepare_image(source, target_dir / f"page-1{source.suffix.lower()}")]

    pages: list[dict] = []
    text_parts: list[str] = []
    for index, image_path in enumerate(page_images, start=1):
        config = " ".join(part for part in [tessdata_config, "--psm 6"] if part)
        text = pytesseract.image_to_string(str(image_path), lang=language, config=config)
        text = _normalize_ocr_text(text)
        text_parts.append(text)
        pages.append({"page": index, "image_relpath": image_path.name, "text": text})

    return OcrResult(
        text="\n\n".join(part for part in text_parts if part).strip(),
        pages=pages,
        engine="tesseract",
        language=language,
        warnings=warnings,
    )


def find_tesseract() -> Path | None:
    path = shutil.which("tesseract")
    if path:
        return Path(path)
    for candidate in COMMON_TESSERACT_PATHS:
        if candidate.exists():
            return candidate
    return None


def get_ocr_status() -> dict:
    tesseract = find_tesseract()
    if not tesseract:
        return {
            "available": False,
            "engine": "tesseract",
            "path": "",
            "languages": [],
            "warnings": ["未找到 tesseract.exe，图片/PDF OCR 暂不可用。"],
        }
    pytesseract.pytesseract.tesseract_cmd = str(tesseract)
    tessdata_config = get_tessdata_config()
    try:
        languages = list_project_languages() or pytesseract.get_languages(config="")
    except Exception as exc:
        return {
            "available": False,
            "engine": "tesseract",
            "path": str(tesseract),
            "languages": [],
            "warnings": [f"Tesseract 语言包读取失败：{exc}"],
        }
    _, warnings = choose_tesseract_language(languages)
    return {
        "available": True,
        "engine": "tesseract",
        "path": str(tesseract),
        "tessdata": str(resolve_tessdata_dir()) if tessdata_config else "",
        "languages": languages,
        "warnings": warnings,
    }


def get_tessdata_config() -> str:
    tessdata_dir = resolve_tessdata_dir()
    if tessdata_dir and any(tessdata_dir.glob("*.traineddata")):
        return f"--tessdata-dir {tessdata_dir}"
    return ""


def list_project_languages() -> list[str]:
    tessdata_dir = resolve_tessdata_dir()
    if not tessdata_dir or not tessdata_dir.exists():
        return []
    return sorted(path.stem for path in tessdata_dir.glob("*.traineddata"))


def resolve_tessdata_dir() -> Path | None:
    if not PROJECT_TESSDATA.exists():
        return None
    RUNTIME_TESSDATA.mkdir(parents=True, exist_ok=True)
    for source in PROJECT_TESSDATA.glob("*.traineddata"):
        destination = RUNTIME_TESSDATA / source.name
        if not destination.exists() or destination.stat().st_size != source.stat().st_size:
            shutil.copy2(source, destination)
    return RUNTIME_TESSDATA


def choose_tesseract_language(languages: list[str]) -> tuple[str, list[str]]:
    available = set(languages)
    warnings: list[str] = []
    if "chi_sim" in available and "eng" in available:
        return "chi_sim+eng", warnings
    if "chi_sim" in available:
        return "chi_sim", warnings
    if "eng" in available:
        warnings.append("当前 Tesseract 未安装 chi_sim 中文语言包，中文 OCR 准确率会很低。")
        return "eng", warnings
    fallback = languages[0] if languages else "eng"
    warnings.append("当前 Tesseract 未安装中文或英文语言包，OCR 可能不可用。")
    return fallback, warnings


def create_ocr_preview(result: OcrResult, output_dir: str | Path) -> Path:
    target_dir = Path(output_dir)
    text_path = target_dir / "ocr_text.txt"
    text_path.write_text(result.text, encoding="utf-8")
    warning_html = "".join(f"<li>{html.escape(item)}</li>" for item in result.warnings)
    pages_html = []
    for page in result.pages:
        image_relpath = html.escape(page.get("image_relpath", ""))
        text = html.escape(page.get("text", ""))
        pages_html.append(
            f"""
            <section class="page">
              <h2>第 {page.get('page')} 页</h2>
              <img src="{image_relpath}" alt="OCR page {page.get('page')}">
              <pre>{text}</pre>
            </section>
            """
        )
    preview = target_dir / "ocr_preview.html"
    preview.write_text(
        f"""
        <!doctype html>
        <html lang="zh-CN">
        <head>
          <meta charset="utf-8">
          <title>OCR 识别结果</title>
          <style>
            body {{ font-family: "Microsoft YaHei", Arial, sans-serif; margin: 24px; color: #17202a; }}
            img {{ max-width: 100%; border: 1px solid #d9dee7; margin: 8px 0 16px; }}
            pre {{ white-space: pre-wrap; line-height: 1.7; background: #f8fafc; padding: 14px; border-radius: 8px; }}
            .warning {{ background: #fffbeb; border: 1px solid #fde68a; padding: 12px; border-radius: 8px; }}
          </style>
        </head>
        <body>
          <h1>OCR 识别结果</h1>
          <p>引擎：{html.escape(result.engine)}；语言：{html.escape(result.language)}</p>
          {f'<div class="warning"><ul>{warning_html}</ul></div>' if warning_html else ''}
          {''.join(pages_html)}
        </body>
        </html>
        """,
        encoding="utf-8",
    )
    return preview


def _render_pdf_pages(source: Path, output_dir: Path) -> list[Path]:
    document = fitz.open(source)
    images: list[Path] = []
    try:
        for index, page in enumerate(document, start=1):
            pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            image_path = output_dir / f"page-{index}.png"
            pixmap.save(image_path)
            images.append(_prepare_image(image_path, image_path))
    finally:
        document.close()
    return images


def _prepare_image(source: Path, destination: Path) -> Path:
    image = Image.open(source)
    try:
        image = ImageOps.exif_transpose(image)
        image = ImageOps.grayscale(image)
        image = ImageOps.autocontrast(image)
        max_side = 2200
        if max(image.size) > max_side:
            image.thumbnail((max_side, max_side), Image.Resampling.LANCZOS)
        destination = destination.with_suffix(".png")
        image.save(destination)
        return destination
    finally:
        image.close()


def _normalize_ocr_text(text: str) -> str:
    lines = [line.strip() for line in text.replace("\r", "\n").splitlines()]
    return "\n".join(line for line in lines if line)
