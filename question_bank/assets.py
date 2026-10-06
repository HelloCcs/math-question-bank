from __future__ import annotations
from pathlib import Path
from urllib.parse import urlparse
from PIL import Image

def safe_asset_path(src: str, data_root: str | Path) -> Path | None:
    root = Path(data_root).resolve()
    value = (urlparse(src).path if "://" in src else src).replace("\\", "/")
    if value.startswith("/data/"): value = value[6:]
    candidate = (root / value.lstrip("/")).resolve()
    try: candidate.relative_to(root)
    except ValueError: return None
    return candidate

def image_kind(path: str | Path, inline_context: bool = False) -> str:
    if inline_context: return "inline-formula"
    try:
        with Image.open(path) as image: width,height=image.size
    except Exception: return "missing"
    if height <= 100 or (height <= 220 and width / max(height,1) >= 1.8): return "inline-formula"
    return "block-figure"

def build_asset_manifest(html: str) -> list[dict]:
    from bs4 import BeautifulSoup
    result=[]
    for node in BeautifulSoup(html or "", "html.parser").find_all("img"):
        src=node.get("src",""); result.append({"src":src,"kind":"inline-formula" if "inline-formula" in (node.get("class") or []) else "block-figure"})
    return result
