from __future__ import annotations

import re
import shutil
import subprocess
import zipfile
from pathlib import Path
import tempfile

from PIL import Image, ImageChops


def convert_docx_to_preview(docx_path: str | Path, output_dir: str | Path) -> tuple[Path, str]:
    source = Path(docx_path)
    target_dir = Path(output_dir)
    target_dir.mkdir(parents=True, exist_ok=True)
    preview_path = target_dir / "preview.htm"

    if _has_native_formulas(source):
        methods = ("pandoc", "word", "libreoffice")
    else:
        methods = ("word", "libreoffice", "pandoc")
    last_error: Exception | None = None
    for method in methods:
        try:
            if method == "word":
                _convert_with_word(source, preview_path)
            elif method == "pandoc":
                _convert_with_pandoc(source, preview_path, target_dir)
            else:
                _convert_with_libreoffice(source, preview_path)
            _finalize_preview_html(preview_path)
            return preview_path, method
        except Exception as exc:
            last_error = exc
    if last_error:
        raise last_error
    raise RuntimeError("No Word preview converter is available")


def _has_native_formulas(source: Path) -> bool:
    try:
        with zipfile.ZipFile(source) as archive:
            document_xml = archive.read("word/document.xml").decode("utf-8", errors="ignore")
    except (OSError, KeyError, zipfile.BadZipFile):
        return False
    return bool(re.search(r"<m:oMath(?:\s|>)|<m:oMathPara(?:\s|>)", document_xml))


def _convert_with_word(source: Path, preview_path: Path) -> None:
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("pywin32 is not available") from exc

    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    shutdown_error: Exception | None = None
    try:
        doc = word.Documents.Open(str(source.resolve()), ReadOnly=True, AddToRecentFiles=False)
        try:
            doc.SaveAs2(str(preview_path.resolve()), FileFormat=10)
        finally:
            try:
                doc.Close(False)
            except Exception as exc:
                shutdown_error = exc
    finally:
        try:
            word.Quit()
        except Exception as exc:
            shutdown_error = shutdown_error or exc
        finally:
            pythoncom.CoUninitialize()
    if shutdown_error is not None:
        raise_if_output_missing(preview_path, shutdown_error)


def raise_if_output_missing(output_path: Path, error: Exception) -> None:
    if not output_path.exists() or output_path.stat().st_size == 0:
        raise error


def _convert_with_libreoffice(source: Path, preview_path: Path) -> None:
    executable = _find_libreoffice()
    if not executable:
        raise RuntimeError("LibreOffice is not available")

    target_dir = preview_path.parent
    before = set(target_dir.glob("*.htm*"))
    command = [
        executable,
        "--headless",
        "--convert-to",
        "html",
        "--outdir",
        str(target_dir),
        str(source.resolve()),
    ]
    subprocess.run(command, check=True, capture_output=True, text=True, timeout=180)

    candidates = [path for path in target_dir.glob("*.htm*") if path not in before]
    candidates.extend(path for path in target_dir.glob(f"{source.stem}.htm*") if path not in candidates)
    if not candidates:
        raise RuntimeError("LibreOffice did not create an HTML preview")
    html_path = max(candidates, key=lambda path: path.stat().st_mtime)
    if html_path.resolve() != preview_path.resolve():
        if preview_path.exists():
            preview_path.unlink()
        html_path.rename(preview_path)
    raise_if_output_missing(preview_path, RuntimeError("LibreOffice HTML preview is empty"))


def _find_libreoffice() -> str | None:
    for command in ("libreoffice", "soffice"):
        executable = shutil.which(command)
        if executable:
            return executable
    for candidate in (
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    ):
        if candidate.exists():
            return str(candidate)
    return None


def _find_pandoc() -> str | None:
    executable = shutil.which("pandoc")
    if executable:
        return executable
    candidates = [
        Path(r"C:\Program Files\Pandoc\pandoc.exe"),
        Path(r"C:\Program Files (x86)\Pandoc\pandoc.exe"),
    ]
    local_appdata = Path.home() / "AppData" / "Local"
    candidates.extend(
        [
            local_appdata / "Pandoc" / "pandoc.exe",
            local_appdata / "Programs" / "Pandoc" / "pandoc.exe",
        ]
    )
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)
    return None


def _convert_with_pandoc(source: Path, preview_path: Path, target_dir: Path) -> None:
    executable = _find_pandoc()
    if not executable:
        raise RuntimeError("Pandoc is not available")
    media_dir = target_dir / "media"
    if media_dir.exists():
        shutil.rmtree(media_dir)
    safe_source = source
    temp_dir: tempfile.TemporaryDirectory[str] | None = None
    if not _is_ascii_path(source):
        temp_dir = tempfile.TemporaryDirectory()
        temp_source = Path(temp_dir.name) / "source.docx"
        shutil.copy2(source, temp_source)
        safe_source = temp_source
    command = [
        executable,
        "--from",
        "docx",
        "--to",
        "html5",
        "--mathjax",
        "--extract-media=.",
        "-o",
        preview_path.name,
        str(safe_source),
    ]
    try:
        subprocess.run(command, check=True, cwd=target_dir)
        _convert_vector_media_with_libreoffice(target_dir, preview_path)
    finally:
        if temp_dir is not None:
            temp_dir.cleanup()


def _convert_vector_media_with_libreoffice(target_dir: Path, preview_path: Path) -> None:
    executable = _find_libreoffice()
    if not executable or not preview_path.exists():
        return

    replacements: dict[str, str] = {}
    for source in list(target_dir.rglob("*.wmf")) + list(target_dir.rglob("*.emf")):
        try:
            subprocess.run(
                [
                    executable,
                    "--headless",
                    "--convert-to",
                    "png",
                    "--outdir",
                    str(source.parent),
                    str(source),
                ],
                check=True,
                capture_output=True,
                text=True,
                timeout=60,
            )
        except Exception:
            continue
        png_path = source.with_suffix(".png")
        if not png_path.exists():
            continue
        _trim_white_margins(png_path)
        old_ref = source.relative_to(target_dir).as_posix()
        new_ref = png_path.relative_to(target_dir).as_posix()
        replacements[old_ref] = new_ref
        replacements[f"./{old_ref}"] = f"./{new_ref}"

    if replacements:
        _rewrite_preview_media_sources(preview_path, replacements)


def _rewrite_preview_media_sources(preview_path: Path, replacements: dict[str, str]) -> None:
    html = preview_path.read_text(encoding="utf-8", errors="ignore")
    for old_ref, new_ref in sorted(replacements.items(), key=lambda item: len(item[0]), reverse=True):
        html = html.replace(old_ref, new_ref)
    preview_path.write_text(html, encoding="utf-8")


def _trim_white_margins(image_path: Path) -> None:
    try:
        with Image.open(image_path) as image:
            rgba = image.convert("RGBA")
            background = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
            diff = ImageChops.difference(rgba, background)
            bbox = diff.getbbox()
            if not bbox:
                return
            cropped = rgba.crop(bbox)
            if cropped.size == rgba.size:
                return
            cropped.save(image_path)
    except Exception:
        return


def _is_ascii_path(path: Path) -> bool:
    try:
        str(path).encode("ascii")
        return True
    except UnicodeEncodeError:
        return False


def _finalize_preview_html(preview_path: Path) -> None:
    html = _read_preview_html(preview_path)
    html = _ensure_utf8_meta(html)
    html = _sanitize_preview_math(html)
    preview_path.write_text(html, encoding="utf-8")


def _sanitize_preview_math(html: str) -> str:
    def replace_overset(match: re.Match[str]) -> str:
        marker = match.group(1)
        body = match.group(2)
        if re.search(r"[A-Za-z0-9\\\\]", marker):
            return match.group(0)
        return rf"\vec{{{body}}}"

    return re.sub(
        r"\\overset\{([^{}]{1,8})\}\{([A-Za-z](?:_[A-Za-z0-9]+)?)\}",
        replace_overset,
        html,
    )


def _read_preview_html(preview_path: Path) -> str:
    raw = preview_path.read_bytes()
    html_ascii = raw.decode("ascii", errors="ignore")
    encodings: list[str] = []
    charset_match = re.search(r"charset\s*=\s*['\"]?([A-Za-z0-9._-]+)", html_ascii, re.IGNORECASE)
    if charset_match:
        encodings.append(charset_match.group(1))
    encodings.extend(["utf-8", "gb18030", "cp936", "utf-16", "latin-1"])
    seen: set[str] = set()
    for encoding in encodings:
        lowered = encoding.lower()
        if lowered in seen:
            continue
        seen.add(lowered)
        try:
            return raw.decode(encoding)
        except (LookupError, UnicodeDecodeError):
            continue
    return raw.decode("utf-8", errors="ignore")


def _ensure_utf8_meta(html: str) -> str:
    if re.search(r"<meta[^>]+charset=", html, re.IGNORECASE):
        html = re.sub(
            r"<meta[^>]+charset=[^>]+>",
            '<meta charset="utf-8">',
            html,
            count=1,
            flags=re.IGNORECASE,
        )
    elif re.search(r"<meta[^>]+http-equiv=['\"]?Content-Type['\"]?[^>]*>", html, re.IGNORECASE):
        html = re.sub(
            r"<meta[^>]+http-equiv=['\"]?Content-Type['\"]?[^>]*>",
            '<meta charset="utf-8">',
            html,
            count=1,
            flags=re.IGNORECASE,
        )
    elif "</head>" in html.lower():
        html = re.sub(r"</head>", '<meta charset="utf-8"></head>', html, count=1, flags=re.IGNORECASE)
    else:
        html = '<meta charset="utf-8">' + html
    return html
