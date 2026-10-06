from __future__ import annotations

import re
import tempfile
import zipfile
from dataclasses import dataclass, replace
from html import unescape
from pathlib import Path
from typing import Sequence

from bs4 import BeautifulSoup, NavigableString, Tag
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont
from .papers import PaperExportModel


HEADERS = ["题号", "题型", "大分类", "小分类", "难度", "来源", "题干", "选项", "答案", "解析"]
INLINE_LATEX_RE = re.compile(r"(\\\((.+?)\\\)|\\\[(.+?)\\\])", re.DOTALL)


@dataclass(frozen=True)
class RunStyle:
    bold: bool = False
    italic: bool = False
    superscript: bool = False
    subscript: bool = False


def export_questions_to_xlsx(questions: Sequence[dict], output_path: str | Path) -> Path:
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "题库"
    sheet.append(HEADERS)
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2F5597")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for question in questions:
        options = question.get("options", "")
        if isinstance(options, list):
            options = "\n".join(options)
        sheet.append(
            [
                question.get("number", ""),
                question.get("question_type", ""),
                question.get("category_major", ""),
                question.get("category_minor", ""),
                question.get("difficulty", ""),
                question.get("source_name", ""),
                question.get("stem", ""),
                options,
                question.get("answer", ""),
                question.get("analysis", ""),
            ]
        )

    widths = [10, 12, 14, 18, 10, 28, 52, 34, 22, 48]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    workbook.save(path)
    return path


QUESTION_TYPE_ORDER = ["单选题", "多选题", "填空题", "解答题"]


def export_questions_to_docx(
    questions: Sequence[dict],
    output_path: str | Path,
    separate_answers: bool = False,
    stems_only: bool = False,
) -> Path:
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    document = Document()
    _set_document_font(document)
    document.add_heading("高中数学题库导出", level=1)

    with tempfile.TemporaryDirectory() as tmpdir:
        asset_dir = Path(tmpdir)
        formula_cache: dict[str, Path] = {}
        
        grouped = _group_questions_for_export(questions)
        for question_type, group in grouped:
            document.add_heading(_section_title(question_type), level=2)
            for number, question in enumerate(group, start=1):
                document.add_paragraph(f"{number}.", style="Heading 3")
                _add_content(document, question.get("stem_html") or "", question.get("stem", ""), asset_dir, formula_cache)
                if not stems_only:
                    _add_content(document, question.get("options_html") or "", question.get("options", ""), asset_dir, formula_cache)
            
            # 如果不分离答案，则按原格式显示答案和解析
        # 第二部分：参考答案（如果需要分离）
        if separate_answers:
            document.add_page_break()
            document.add_heading("参考答案", level=1)
            
            for index, question in enumerate(questions, start=1):
                question_number = question.get('number') or str(index)
                has_content = False
                
                # 检查是否有答案或解析
                if question.get("answer") or question.get("answer_html") or question.get("analysis") or question.get("analysis_html"):
                    has_content = True
                
                if not has_content:
                    continue
                
                # 添加题号
                document.add_paragraph(f"{question_number}.", style="Heading 3")
                
                # 添加答案
                answer = question.get("answer") or ""
                if answer or question.get("answer_html"):
                    _add_labeled_content(
                        document,
                        "答案：",
                        question.get("answer_html") or "",
                        answer,
                        asset_dir,
                        formula_cache,
                    )
                
                # 添加解析
                analysis = question.get("analysis") or ""
                if analysis or question.get("analysis_html"):
                    _add_labeled_content(
                        document,
                        "解析：",
                        question.get("analysis_html") or "",
                        analysis,
                        asset_dir,
                        formula_cache,
                    )

        document.save(path)
    return path


def _group_questions_for_export(questions: Sequence[dict]) -> list[tuple[str, list[dict]]]:
    grouped: dict[str, list[dict]] = {question_type: [] for question_type in QUESTION_TYPE_ORDER}
    other: list[dict] = []
    for question in questions:
        question_type = str(question.get("question_type") or "未分类")
        if question_type in grouped:
            grouped[question_type].append(question)
        else:
            other.append(question)
    result = [(question_type, rows) for question_type, rows in grouped.items() if rows]
    if other:
        result.append(("未分类", other))
    return result


def _section_title(question_type: str) -> str:
    names = {"单选题": "一、单选题", "多选题": "二、多选题", "填空题": "三、填空题", "解答题": "四、解答题"}
    return names.get(question_type, question_type)


def export_paper_docx(model: PaperExportModel, output_path: str | Path, *, answers: bool = False) -> Path:
    path=Path(output_path); path.parent.mkdir(parents=True,exist_ok=True); document=Document(); _set_document_font(document)
    document.add_heading(model.title + (" - 答案与解析" if answers else ""), level=1)
    if not answers:
        meta="  ".join(item for item in [model.school,model.grade,model.class_name,f"考试时间：{model.exam_time}",f"满分：{model.total_score}"] if item)
        if meta: document.add_paragraph(meta)
        if model.instructions: document.add_paragraph(model.instructions)
    with tempfile.TemporaryDirectory() as tmpdir:
        assets=Path(tmpdir); cache={}; current_type=""
        for q in model.questions:
            qtype=q.get("question_type") or "未分类"
            if qtype!=current_type: document.add_heading(qtype,level=2); current_type=qtype
            document.add_paragraph(f"{q['export_number']}.",style="Heading 3")
            _add_content(document,q.get("stem_html") or "",q.get("stem", ""),assets,cache)
            _add_content(document,q.get("options_html") or "",q.get("options", ""),assets,cache)
            if answers:
                _add_labeled_content(document,"答案：",q.get("answer_html") or "",q.get("answer","") or "暂无独立答案",assets,cache)
                _add_labeled_content(document,"解析：",q.get("analysis_html") or "",q.get("analysis","") or "暂无解析",assets,cache)
        document.save(path)
    return path


def export_paper_word_zip(model: PaperExportModel, output_path: str | Path) -> Path:
    path=Path(output_path); path.parent.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        root=Path(tmp); question=export_paper_docx(model,root/"试卷-题目.docx"); answer=export_paper_docx(model,root/"试卷-答案与解析.docx",answers=True)
        with zipfile.ZipFile(path,"w",zipfile.ZIP_DEFLATED) as archive:
            archive.write(question,question.name); archive.write(answer,answer.name)
    return path


def export_paper_pdf(model: PaperExportModel, output_path: str | Path) -> Path:
    path=Path(output_path); path.parent.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        docx=export_paper_docx(model,Path(tmp)/"试卷-题目.docx")
        try:
            _convert_docx_to_pdf_with_word(docx,path)
        except Exception:
            _export_questions_to_simple_pdf(
                model.questions,
                path,
                title=model.title,
                meta="  ".join(item for item in [model.school, model.grade, model.class_name, model.exam_time, f"{model.total_score} \u5206"] if item),
            )
    return path


def export_questions_to_pdf(questions: Sequence[dict], output_path: str | Path, stems_only: bool = False) -> Path:
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = Path(tmpdir) / "questions.docx"
        export_questions_to_docx(questions, docx_path, stems_only=stems_only)
        try:
            _convert_docx_to_pdf_with_word(docx_path, path)
        except Exception:
            _export_questions_to_simple_pdf(questions, path, title="\u9898\u5e93\u5bfc\u51fa", include_answers=not stems_only)
    return path


def export_questions_to_image_zip(questions: Sequence[dict], output_path: str | Path) -> Path:
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        image_paths: list[tuple[Path, str]] = []
        used_names: set[str] = set()
        for index, question in enumerate(questions, start=1):
            label = question.get("number") or str(index)
            image_name = _unique_image_name(index, str(label), used_names)
            image_path = tmp_path / image_name
            _render_question_png(question, image_path)
            image_paths.append((image_path, image_name))
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
            for image_path, image_name in image_paths:
                archive.write(image_path, image_name)
    return path


def _set_document_font(document: Document) -> None:
    _set_style_font(document, "Normal", "宋体", 10)
    _set_style_font(document, "Heading 1", "黑体", 14)
    _set_style_font(document, "Heading 2", "黑体", 12)
    _set_style_font(document, "Heading 3", "黑体", 10.5)


def _set_style_font(document: Document, style_name: str, font_name: str, size: float) -> None:
    style = document.styles[style_name]
    style.font.name = font_name
    style.font.size = Pt(size)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)


def _add_content(
    document: Document,
    html: str,
    fallback: str | list[str],
    asset_dir: Path,
    formula_cache: dict[str, Path],
) -> None:
    if html:
        soup = BeautifulSoup(html, "html.parser")
        blocks = _top_level_content_blocks(soup)
        for block in blocks:
            if isinstance(block, Tag) and block.name == "table":
                _add_html_table(document, block, asset_dir, formula_cache)
            elif _looks_like_compact_option_block(block):
                _add_compact_option_block(document, block, asset_dir, formula_cache)
            else:
                paragraph = document.add_paragraph()
                _append_html_fragment(paragraph, block, asset_dir, formula_cache, RunStyle())
        return

    if isinstance(fallback, list):
        lines = [line for line in fallback if line]
    else:
        lines = [line for line in str(fallback or "").splitlines() if line]
    for line in lines or [""]:
        if line:
            document.add_paragraph(line)


def _add_labeled_content(
    document: Document,
    label: str,
    html: str,
    fallback: str,
    asset_dir: Path,
    formula_cache: dict[str, Path],
) -> None:
    paragraph = document.add_paragraph()
    paragraph.add_run(label).bold = True
    if html:
        soup = BeautifulSoup(html, "html.parser")
        blocks = _top_level_content_blocks(soup)
        first = True
        for block in blocks:
            if isinstance(block, Tag) and block.name == "table":
                if first:
                    first = False
                _add_html_table(document, block, asset_dir, formula_cache)
                continue
            active = paragraph if first else document.add_paragraph()
            _append_html_fragment(active, block, asset_dir, formula_cache, RunStyle())
            first = False
    elif fallback:
        paragraph.add_run(fallback)


def _looks_like_compact_option_block(block: Tag) -> bool:
    if not isinstance(block, Tag) or block.name != "p":
        return False
    text = re.sub(r"\s+", " ", block.get_text(" ", strip=True).replace("\xa0", " ")).strip()
    return len(re.findall(r"(?<!\w)[A-H][\.\uff0e]?", text)) >= 3


def _split_compact_options(text: str) -> list[tuple[str, str]]:
    cleaned = re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()
    matches = list(re.finditer(r"(?<!\w)([A-H])[\.\uff0e]?", cleaned))
    if len(matches) < 2:
        return []
    items: list[tuple[str, str]] = []
    for index, match in enumerate(matches):
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(cleaned)
        content = cleaned[start:end].strip()
        content = re.sub(r"^[\s\.\uff0e、:：]+", "", content)
        if content:
            items.append((match.group(1), content))
    return items


def _flatten_option_tokens(node) -> list[tuple[str, str]]:
    tokens: list[tuple[str, str]] = []
    if isinstance(node, NavigableString):
        tokens.extend(("text", char) for char in str(node))
    elif isinstance(node, Tag):
        if node.name == "img":
            tokens.append(("html", str(node)))
        else:
            for child in node.children:
                tokens.extend(_flatten_option_tokens(child))
    return tokens


def _split_compact_option_fragments(block: Tag) -> list[tuple[str, str]]:
    tokens: list[tuple[str, str]] = []
    for child in block.children:
        tokens.extend(_flatten_option_tokens(child))

    markers: list[tuple[int, str, int]] = []
    for index, (kind, value) in enumerate(tokens):
        if kind != "text" or value not in "ABCDEFGH":
            continue
        previous = tokens[index - 1][1] if index else ""
        if previous and not previous.isspace():
            continue
        next_index = index + 1
        while next_index < len(tokens) and tokens[next_index][0] == "text" and tokens[next_index][1].isspace():
            next_index += 1
        if next_index < len(tokens) and tokens[next_index][0] == "text" and tokens[next_index][1] in ".．、:：":
            next_index += 1
        markers.append((index, value, next_index))

    if len(markers) < 3:
        return []

    fragments: list[tuple[str, str]] = []
    for marker_index, (start, label, content_start) in enumerate(markers):
        end = markers[marker_index + 1][0] if marker_index + 1 < len(markers) else len(tokens)
        content = "".join(value for kind, value in tokens[content_start:end])
        content = re.sub(r"^[\s\.\uff0e、:：]+", "", content)
        content = re.sub(r"\s+", " ", content).strip()
        if content:
            fragments.append((label, content))
    return fragments


def _format_option_paragraph(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.left_indent = Inches(0.28)
    paragraph.paragraph_format.first_line_indent = Inches(-0.18)


def _add_compact_option_block(
    document: Document,
    block: Tag,
    asset_dir: Path,
    formula_cache: dict[str, Path],
) -> None:
    options = _split_compact_option_fragments(block)
    if not options:
        options = _split_compact_options(block.get_text(" ", strip=True))
    if not options:
        paragraph = document.add_paragraph()
        _format_option_paragraph(paragraph)
        _append_html_fragment(paragraph, block, asset_dir, formula_cache, RunStyle())
        return

    for label, content in options:
        paragraph = document.add_paragraph()
        _format_option_paragraph(paragraph)
        run = paragraph.add_run(f"{label}. ")
        run.bold = True
        content_soup = BeautifulSoup(content, "html.parser")
        for child in content_soup.contents:
            _append_html_fragment(paragraph, child, asset_dir, formula_cache, RunStyle())


def _top_level_content_blocks(soup: BeautifulSoup) -> list:
    blocks = []
    for child in soup.children:
        if not isinstance(child, Tag):
            continue
        if child.name == "p" and child.find("table", recursive=False):
            blocks.extend(
                item for item in child.children
                if isinstance(item, Tag) and item.name in {"p", "table"}
            )
            continue
        if child.name in {"p", "table"}:
            blocks.append(child)
    return _restore_flattened_score_tables(blocks) or [soup]


def _restore_flattened_score_tables(blocks: list[Tag]) -> list[Tag]:
    result: list[Tag] = []
    index = 0
    while index < len(blocks):
        restored = _collect_flattened_score_table(blocks, index)
        if restored:
            table, next_index = restored
            result.append(table)
            index = next_index
            continue
        result.append(blocks[index])
        index += 1
    return result


def _collect_flattened_score_table(blocks: list[Tag], start: int) -> tuple[Tag, int] | None:
    labels = {"原始分", "转换分", "人数"}
    ordered_rows: list[tuple[str, list[str]]] = []
    index = start

    while index < len(blocks):
        label = _block_text(blocks[index])
        if label not in labels:
            break
        index += 1
        values: list[str] = []
        while index < len(blocks):
            text = _block_text(blocks[index])
            if text in labels:
                break
            if not _looks_like_table_value(text):
                break
            values.append(text)
            index += 1
        if not values:
            return None
        ordered_rows.append((label, values))

    if len(ordered_rows) < 2:
        return None

    soup = BeautifulSoup("", "html.parser")
    table = soup.new_tag("table")
    max_values = max(len(values) for _, values in ordered_rows)
    for label, values in ordered_rows:
        tr = soup.new_tag("tr")
        for value in [label, *values, *[""] * (max_values - len(values))]:
            td = soup.new_tag("td")
            td.string = value
            tr.append(td)
        table.append(tr)
    return table, index


def _block_text(block: Tag) -> str:
    return re.sub(r"\s+", "", block.get_text(" ", strip=True).replace("\xa0", " "))


def _looks_like_table_value(text: str) -> bool:
    if not text:
        return False
    return bool(re.fullmatch(r"[\d.％%分+-]+", text))


def _add_html_table(
    document: Document,
    table_node: Tag,
    asset_dir: Path,
    formula_cache: dict[str, Path],
) -> None:
    rows = []
    for row_node in table_node.find_all("tr"):
        cells = row_node.find_all(["td", "th"], recursive=False)
        if cells:
            rows.append(cells)
    if not rows:
        return

    column_count = max(len(row) for row in rows)
    table = document.add_table(rows=len(rows), cols=column_count)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = True

    for row_index, cells in enumerate(rows):
        for column_index in range(column_count):
            target = table.cell(row_index, column_index)
            target.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            _set_cell_margins(target, top=80, left=100, bottom=80, right=100)
            paragraph = target.paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            if column_index >= len(cells):
                continue
            cell_node = cells[column_index]
            cell_blocks = cell_node.find_all("p", recursive=False) or [cell_node]
            for block_index, block in enumerate(cell_blocks):
                active = paragraph if block_index == 0 else target.add_paragraph()
                active.alignment = WD_ALIGN_PARAGRAPH.CENTER
                active.paragraph_format.space_before = Pt(0)
                active.paragraph_format.space_after = Pt(0)
                _append_html_fragment(active, block, asset_dir, formula_cache, RunStyle())


def _set_cell_margins(cell, *, top: int, left: int, bottom: int, right: int) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for name, value in {"top": top, "left": left, "bottom": bottom, "right": right}.items():
        node = tc_mar.find(qn(f"w:{name}"))
        if node is None:
            node = OxmlElement(f"w:{name}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def _append_html_fragment(
    paragraph,
    node,
    asset_dir: Path,
    formula_cache: dict[str, Path],
    style: RunStyle,
) -> None:
    if isinstance(node, NavigableString):
        _append_text(paragraph, str(node), style, asset_dir, formula_cache)
        return
    if not isinstance(node, Tag):
        return

    if node.name == "br":
        paragraph.add_run().add_break()
        return
    if node.name == "img":
        image_path = _resolve_image_src(node.get("src", ""))
        if image_path and image_path.exists():
            run = paragraph.add_run()
            run.add_picture(str(image_path), **_image_size_args(image_path, node.get("class") or []))
        else:
            _apply_run_style(paragraph.add_run("[图片]"), style)
        return
    if node.name == "table":
        _apply_run_style(paragraph.add_run("[表格]"), style)
        return

    next_style = style
    if node.name in {"strong", "b"}:
        next_style = replace(next_style, bold=True)
    elif node.name in {"em", "i"}:
        next_style = replace(next_style, italic=True)
    elif node.name == "sup":
        next_style = replace(next_style, superscript=True)
    elif node.name == "sub":
        next_style = replace(next_style, subscript=True)

    for child in node.children:
        _append_html_fragment(paragraph, child, asset_dir, formula_cache, next_style)


def _append_text(
    paragraph,
    text: str,
    style: RunStyle,
    asset_dir: Path,
    formula_cache: dict[str, Path],
) -> None:
    plain = unescape(text or "")
    if not plain:
        return
    position = 0
    for match in INLINE_LATEX_RE.finditer(plain):
        prefix = plain[position : match.start()]
        if prefix:
            _apply_run_style(paragraph.add_run(prefix), style)
        formula = (match.group(2) or match.group(3) or "").strip()
        if formula and _append_formula_image(paragraph, formula, asset_dir, formula_cache, style):
            position = match.end()
            continue
        _apply_run_style(paragraph.add_run(_formula_to_readable_text(formula or match.group(0))), style)
        position = match.end()

    suffix = plain[position:]
    if suffix:
        _apply_run_style(paragraph.add_run(suffix), style)


def _append_formula_image(
    paragraph,
    formula: str,
    asset_dir: Path,
    formula_cache: dict[str, Path],
    style: RunStyle,
) -> bool:
    image_path = _render_formula_image(formula, asset_dir, formula_cache)
    if not image_path:
        return False
    run = paragraph.add_run()
    # 字体大小11pt，公式图片高度最多为字体的1.5倍（16.5pt）
    font_size = 11
    max_height = font_size * 1.5
    height = Pt(max_height if not style.superscript and not style.subscript else max_height * 0.8)
    run.add_picture(str(image_path), height=height)
    return True


def _render_formula_image(formula: str, asset_dir: Path, formula_cache: dict[str, Path]) -> Path | None:
    cached = formula_cache.get(formula)
    if cached and cached.exists():
        return cached
    try:
        import matplotlib

        matplotlib.use("Agg")
        from matplotlib import pyplot as plt
    except Exception:
        return None

    image_path = asset_dir / f"formula-{len(formula_cache) + 1}.png"
    fig = plt.figure(figsize=(0.01, 0.01))
    try:
        fig.patch.set_alpha(0.0)
        fig.text(0, 0, f"${formula}$", fontsize=14, color="black")
        fig.savefig(
            image_path,
            dpi=220,
            bbox_inches="tight",
            pad_inches=0.04,
            transparent=True,
        )
    except Exception:
        plt.close(fig)
        return None
    plt.close(fig)
    if not image_path.exists() or image_path.stat().st_size == 0:
        return None
    formula_cache[formula] = image_path
    return image_path


def _formula_to_readable_text(formula: str) -> str:
    text = formula.strip()
    greek = {
        r"\alpha": "α",
        r"\beta": "β",
        r"\gamma": "γ",
        r"\theta": "θ",
        r"\lambda": "λ",
        r"\phi": "φ",
        r"\pi": "π",
        r"\infty": "∞",
        r"\forall": "∀",
        r"\exists": "∃",
        r"\in": "∈",
        r"\cup": "∪",
        r"\cap": "∩",
        r"\geq": "≥",
        r"\leq": "≤",
        r"\neq": "≠",
        r"\times": "×",
        r"\cdot": "·",
        r"\pm": "±",
    }
    for raw, target in greek.items():
        text = text.replace(raw, target)

    text = re.sub(r"\\frac\{([^{}]+)\}\{([^{}]+)\}", r"(\1/\2)", text)
    text = re.sub(r"\\sqrt\{([^{}]+)\}", r"√(\1)", text)
    text = re.sub(r"\\log_\{([^{}]+)\}", r"log_\1", text)
    text = re.sub(r"\\[lr]brack", "", text)
    text = text.replace(r"\left", "").replace(r"\right", "")
    text = text.replace("{", "").replace("}", "")
    text = text.replace("\\", "")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _apply_run_style(run, style: RunStyle) -> None:
    run.bold = style.bold
    run.italic = style.italic
    run.font.superscript = style.superscript
    run.font.subscript = style.subscript


def _resolve_image_src(src: str) -> Path | None:
    if not src:
        return None
    clean_src = src.replace("\\", "/")
    if clean_src.startswith("/data/"):
        return Path.cwd() / clean_src.lstrip("/")
    candidate = Path(clean_src)
    if candidate.is_absolute():
        return candidate
    return Path.cwd() / clean_src


def _image_size_args(image_path: Path, classes: Sequence[str] | None = None) -> dict:
    try:
        with Image.open(image_path) as image:
            width, height = image.size
    except Exception:
        return {"height": Pt(16)}

    class_set = {str(item) for item in (classes or [])}
    aspect_ratio = width / height if height else 1
    is_block_figure = "block-figure" in class_set
    is_formula = "inline-formula" in class_set and not is_block_figure

    if is_formula:
        return {"height": Pt(15 if height <= 22 else 17)}

    width_inches = width / 96
    height_inches = height / 96
    max_width_inches = 5.8
    max_height_inches = 3.6

    if is_block_figure:
        min_width_inches = 2.4 if aspect_ratio >= 1 else 1.8
        width_inches = max(width_inches, min_width_inches)
        width_inches = min(width_inches, max_width_inches)
        return {"width": Inches(width_inches)}

    if width_inches > max_width_inches:
        return {"width": Inches(max_width_inches)}
    if height_inches > max_height_inches:
        return {"height": Inches(max_height_inches)}
    return {"width": Inches(width_inches)}


def _unique_image_name(index: int, label: str, used_names: set[str]) -> str:
    safe_label = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff_-]+", "-", label).strip("-") or str(index)
    base = f"question-{safe_label}.png"
    if base not in used_names:
        used_names.add(base)
        return base
    name = f"question-{index}-{safe_label}.png"
    counter = 2
    while name in used_names:
        name = f"question-{index}-{safe_label}-{counter}.png"
        counter += 1
    used_names.add(name)
    return name


def _export_questions_to_simple_pdf(
    questions: Sequence[dict],
    output_path: Path,
    *,
    title: str,
    meta: str = "",
    include_answers: bool = False,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    page_width, page_height = 1240, 1754
    margin = 72
    body_font = _load_font(28)
    title_font = _load_font(38)
    small_font = _load_font(24)
    line_height = 44
    pages: list[Image.Image] = []
    page = Image.new("RGB", (page_width, page_height), "white")
    draw = ImageDraw.Draw(page)
    y = margin

    def new_page() -> None:
        nonlocal page, draw, y
        pages.append(page)
        page = Image.new("RGB", (page_width, page_height), "white")
        draw = ImageDraw.Draw(page)
        y = margin

    def ensure_space(height: int) -> None:
        if y + height > page_height - margin:
            new_page()

    def add_text(text: str, font: ImageFont.ImageFont = body_font, *, gap: int = 8) -> None:
        nonlocal y
        for line in _wrap_text_for_pdf(text, font, page_width - margin * 2):
            ensure_space(line_height)
            draw.text((margin, y), line, fill="#111827", font=font)
            y += line_height
        y += gap

    def add_image(image_path: Path) -> None:
        nonlocal y
        try:
            with Image.open(image_path) as source:
                image = source.convert("RGB")
        except Exception:
            return
        max_width = page_width - margin * 2
        max_height = 420
        scale = min(max_width / image.width, max_height / image.height, 1.8)
        width = max(1, int(image.width * scale))
        height = max(1, int(image.height * scale))
        ensure_space(height + 20)
        resized = image.resize((width, height), Image.Resampling.LANCZOS)
        page.paste(resized, (margin, y))
        y += height + 20

    if title:
        add_text(title, title_font, gap=18)
    if meta:
        add_text(meta, small_font, gap=24)

    for index, question in enumerate(questions, start=1):
        number = question.get("export_number") or question.get("number") or index
        qtype = str(question.get("question_type") or "")
        add_text(f"{number}. {qtype}".strip(), body_font, gap=6)
        for line in _question_plain_lines(question, include_answers=include_answers):
            add_text(line, body_font, gap=4)
        for image_path in _question_block_images(question):
            add_image(image_path)
        y += 18

    pages.append(page)
    pages[0].save(output_path, "PDF", resolution=150.0, save_all=True, append_images=pages[1:])


def _question_plain_lines(question: dict, *, include_answers: bool) -> list[str]:
    sections = [
        _html_text(question.get("stem_html") or "") or str(question.get("stem") or ""),
        _html_text(question.get("options_html") or "") or _stringify_lines(question.get("options") or ""),
    ]
    if include_answers:
        answer = _html_text(question.get("answer_html") or "") or str(question.get("answer") or "")
        analysis = _html_text(question.get("analysis_html") or "") or str(question.get("analysis") or "")
        if answer:
            sections.append(f"\u7b54\u6848\uff1a{answer}")
        if analysis:
            sections.append(f"\u89e3\u6790\uff1a{analysis}")
    return [re.sub(r"\s+", " ", item).strip() for item in sections if str(item).strip()]


def _stringify_lines(value) -> str:
    if isinstance(value, list):
        return "\n".join(str(item) for item in value if str(item).strip())
    return str(value or "")


def _html_text(html: str) -> str:
    if not html:
        return ""
    soup = BeautifulSoup(html, "html.parser")
    for image in soup.find_all("img"):
        image.replace_with(" ")
    return soup.get_text(" ", strip=True)


def _question_block_images(question: dict) -> list[Path]:
    paths: list[Path] = []
    for html in (question.get("stem_html") or "", question.get("options_html") or ""):
        if not html:
            continue
        soup = BeautifulSoup(html, "html.parser")
        for image in soup.find_all("img"):
            classes = {str(item) for item in (image.get("class") or [])}
            if "inline-formula" in classes and "block-figure" not in classes:
                continue
            image_path = _resolve_image_src(image.get("src", ""))
            if image_path and image_path.exists():
                paths.append(image_path)
    return paths


def _wrap_text_for_pdf(text: str, font: ImageFont.ImageFont, max_width: int) -> list[str]:
    lines: list[str] = []
    for raw_line in str(text or "").splitlines() or [""]:
        lines.extend(_wrap_text(raw_line, font, max_width))
    return lines or [""]


def _convert_docx_to_pdf_with_word(docx_path: Path, output_path: Path) -> None:
    docx_path = Path(docx_path).resolve()
    output_path = Path(output_path).resolve()
    word_error: Exception | None = None
    try:
        _convert_docx_to_pdf_with_libreoffice(docx_path, output_path)
        return
    except Exception as libreoffice_error:
        word_error = libreoffice_error

    if not _word_is_available():
        raise RuntimeError(
            "PDF 导出需要安装 Microsoft Word 或 LibreOffice。当前未检测到可用转换器；"
            "WPS 可以打开 Word，但通常不能被本系统后台调用生成 PDF。"
        ) from word_error

    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        _convert_docx_to_pdf_with_libreoffice(docx_path, output_path, exc)
        return

    try:
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        shutdown_error: Exception | None = None
        try:
            doc = word.Documents.Open(str(docx_path), ReadOnly=True, AddToRecentFiles=False)
            try:
                doc.ExportAsFixedFormat(str(output_path), 17)
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
            pythoncom.CoUninitialize()
        if shutdown_error is not None and (not output_path.exists() or output_path.stat().st_size == 0):
            raise shutdown_error
        if output_path.exists() and output_path.stat().st_size:
            return
    except Exception as exc:
        word_error = exc
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    try:
        _convert_docx_to_pdf_with_libreoffice(docx_path, output_path, word_error)
    except Exception:
        if word_error is not None:
            raise word_error
        raise


def _convert_docx_to_pdf_with_libreoffice(
    docx_path: Path,
    output_path: Path,
    original_error: Exception | None = None,
) -> None:
    import shutil
    import subprocess

    binary = shutil.which("libreoffice") or shutil.which("soffice")
    if not binary:
        raise RuntimeError("当前环境缺少 pywin32 或 LibreOffice，无法导出 PDF。") from original_error

    output_path.parent.mkdir(parents=True, exist_ok=True)
    command = [
        binary,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(output_path.parent),
        str(docx_path),
    ]
    subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
    generated = output_path.parent / f"{docx_path.stem}.pdf"
    if generated.exists() and generated != output_path:
        generated.replace(output_path)
    if not output_path.exists() or output_path.stat().st_size == 0:
        raise RuntimeError("LibreOffice 未生成有效 PDF 文件。")


def _word_is_available() -> bool:
    import shutil

    if shutil.which("WINWORD.EXE"):
        return True
    candidates = [
        Path(r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"),
        Path(r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"),
        Path(r"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE"),
        Path(r"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE"),
        Path(r"C:\Program Files\Microsoft Office\Office15\WINWORD.EXE"),
        Path(r"C:\Program Files (x86)\Microsoft Office\Office15\WINWORD.EXE"),
    ]
    if any(path.exists() for path in candidates):
        return True
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE"):
            return True
    except Exception:
        return False


def _render_question_png(question: dict, output_path: Path) -> None:
    font = _load_font(26)
    small_font = _load_font(22)
    width = 1200
    margin = 48
    lines = [f"{question.get('number', '')}. {question.get('stem', '')}".strip()]
    options = question.get("options") or ""
    if isinstance(options, list):
        options = "\n".join(options)
    if options:
        lines.extend(options.splitlines())
    if question.get("answer"):
        lines.append(f"答案：{question.get('answer')}")
    if question.get("analysis"):
        lines.append(f"解析：{question.get('analysis')}")

    wrapped: list[tuple[str, ImageFont.FreeTypeFont | ImageFont.ImageFont]] = []
    for line in lines:
        active_font = font if not line.startswith(("答案：", "解析：")) else small_font
        wrapped.extend((item, active_font) for item in _wrap_text(line, active_font, width - margin * 2))
        wrapped.append(("", active_font))

    line_height = 38
    height = max(260, margin * 2 + line_height * len(wrapped))
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    y = margin
    for line, line_font in wrapped:
        if line:
            draw.text((margin, y), line, fill="#1f2937", font=line_font)
        y += line_height
    image.save(output_path)


def _wrap_text(text: str, font: ImageFont.ImageFont, max_width: int) -> list[str]:
    if not text:
        return [""]
    lines: list[str] = []
    current = ""
    image = Image.new("RGB", (10, 10))
    draw = ImageDraw.Draw(image)
    for char in text:
        trial = current + char
        bbox = draw.textbbox((0, 0), trial, font=font)
        if bbox[2] - bbox[0] <= max_width or not current:
            current = trial
        else:
            lines.append(current)
            current = char
    if current:
        lines.append(current)
    return lines


def _load_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        Path("C:/Windows/Fonts/msyh.ttc"),
        Path("C:/Windows/Fonts/simsun.ttc"),
        Path("C:/Windows/Fonts/simhei.ttf"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return ImageFont.truetype(str(candidate), size=size)
    return ImageFont.load_default()
