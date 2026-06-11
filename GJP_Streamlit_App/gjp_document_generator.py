"""
GJP document generator page for Streamlit.

What it does
------------
1) Uploads the level-differentiated KSA Excel file.
2) Uploads one GJP Word template.
3) Optionally uploads a Word file containing Work interaction and Introduction text.
4) Generates 10 Word documents: P-1, P-2, P-3, P-4, P-5, D-1, NO-A, NO-B, NO-C, NO-D.
5) Replaces level text, updates work experience years, replaces Work Interactions and Introduction,
   builds Responsibilities and Skills/KSA tables, and returns a downloadable ZIP.

Install requirements
--------------------
pip install streamlit python-docx openpyxl

How to use in your main GJP_website.py
--------------------------------------
from gjp_document_generator import render_gjp_document_generator
render_gjp_document_generator()
"""

from __future__ import annotations

import io
import re
import zipfile
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

try:
    import streamlit as st
except ModuleNotFoundError:
    st = None
from docx import Document
from docx.document import Document as DocumentObject
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


# -----------------------------------------------------------------------------
# Configuration
# -----------------------------------------------------------------------------

OUTPUT_LEVELS = ["P-1", "P-2", "P-3", "P-4", "P-5", "D-1", "NO-A", "NO-B", "NO-C", "NO-D"]

# NO levels usually mirror the P-level progression for responsibilities/KSAs.
LEVEL_EQUIVALENT = {
    "NO-A": "P-1",
    "NO-B": "P-2",
    "NO-C": "P-3",
    "NO-D": "P-4",
}

# Work experience wording requested by the user.
WORK_EXPERIENCE_YEARS = {
    "P-1": None,      # keep the template wording
    "NO-A": None,     # keep the template wording
    "P-2": "two",
    "NO-B": "two",
    "P-3": "five",
    "NO-C": "five",
    "P-4": "seven",
    "NO-D": "seven",
    "P-5": "ten",
    "D-1": "fifteen",
}

# Theme colors learned from the uploaded 6.2 Transposing KSA mapping template.
# The colors are used in both the Responsibilities table and the KSA mapping table.
THEME_COLORS = {
    "Risk Assessment & Security Planning": "D9D9D9",
    "Stakeholder Engagement & Coordination": "FFF099",
    "Crisis Management & Response": "A9D08E",
    "Security Operations & Compliance": "F8CBAD",
    "Capacity Building & Training": "AEAAAA",
    "Reporting & Information Management": "CC99FF",
    "Peace Operations & Mission Support": "92D050",
    "Managerial": "C65911",
}
DEFAULT_THEME_COLOR = "D9D9D9"
HEADER_FILL = "000000"
HEADER_FONT = "FFFFFF"


@dataclass
class Responsibility:
    number: int
    theme: str
    text: str
    source_code: str = ""
    mandatory: bool = False


@dataclass
class KsaRow:
    number: int
    text: str
    responsibilities: List[int] = field(default_factory=list)


@dataclass
class LevelTables:
    level: str
    specialty: str
    responsibilities: List[Responsibility]
    ksa_rows: List[KsaRow]


# -----------------------------------------------------------------------------
# Basic Word helpers
# -----------------------------------------------------------------------------

def _normalize_level(level: str) -> str:
    s = str(level).strip().upper().replace("–", "-").replace("—", "-")
    # Convert compact forms from Excel, such as P2/D1/NOA, to P-2/D-1/NO-A.
    m = re.fullmatch(r"P\s*-?\s*([1-5])", s)
    if m:
        return f"P-{m.group(1)}"
    m = re.fullmatch(r"D\s*-?\s*1", s)
    if m:
        return "D-1"
    m = re.fullmatch(r"NO\s*-?\s*([A-D])", s)
    if m:
        return f"NO-{m.group(1)}"
    return s


def _compact_level(level: str) -> str:
    return _normalize_level(level).replace("-", "")


def _para_text(paragraph) -> str:
    return "".join(run.text for run in paragraph.runs) if paragraph.runs else paragraph.text


def _replace_in_paragraph(paragraph, replacements: Dict[str, str]) -> None:
    """Replace text in a paragraph while keeping paragraph style. Run-level style may be flattened."""
    text = _para_text(paragraph)
    new_text = text
    for old, new in replacements.items():
        new_text = new_text.replace(old, new)
    if new_text != text:
        # Preserve the style of the first run when possible.
        if paragraph.runs:
            first_run = paragraph.runs[0]
            for r in list(paragraph.runs):
                r.text = ""
            first_run.text = new_text
        else:
            paragraph.add_run(new_text)


def replace_level_everywhere(doc: DocumentObject, source_level: str, target_level: str) -> None:
    """Replace the template level everywhere: body paragraphs, tables, headers and footers."""
    levels_to_replace = {source_level, _normalize_level(source_level), _compact_level(source_level)}
    replacements = {}
    for old in levels_to_replace:
        replacements[old] = target_level if "-" in old else _compact_level(target_level)

    containers = [doc]
    for section in doc.sections:
        containers.extend([section.header, section.footer])

    for container in containers:
        for paragraph in container.paragraphs:
            _replace_in_paragraph(paragraph, replacements)
        for table in container.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        _replace_in_paragraph(paragraph, replacements)


def detect_template_level(doc: DocumentObject) -> str:
    text = "\n".join(p.text for p in doc.paragraphs[:20])
    m = re.search(r"\b(P-\d|D-1|NO-[A-D])\b", text, flags=re.I)
    return _normalize_level(m.group(1)) if m else "P-2"


def replace_work_experience_years(doc: DocumentObject, target_level: str) -> None:
    year_word = WORK_EXPERIENCE_YEARS.get(target_level)
    if not year_word:
        return

    pattern = re.compile(r"(A minimum of\s+)(one|two|three|four|five|six|seven|eight|nine|ten|fifteen|\d+)(\s+years)", re.I)

    def replace_years_in_paragraph(paragraph) -> None:
        text = _para_text(paragraph)
        new_text = pattern.sub(lambda m: f"{m.group(1)}{year_word}{m.group(3)}", text)
        if new_text != text:
            if paragraph.runs:
                first = paragraph.runs[0]
                for r in paragraph.runs:
                    r.text = ""
                first.text = new_text
            else:
                paragraph.add_run(new_text)

    for paragraph in doc.paragraphs:
        replace_years_in_paragraph(paragraph)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_years_in_paragraph(paragraph)


# -----------------------------------------------------------------------------
# Parse Work interaction / Introduction Word file
# -----------------------------------------------------------------------------

def parse_intro_interaction_docx(path: str | Path) -> Tuple[Dict[str, str], Dict[str, str]]:
    """Return two dictionaries: interactions[level] and introductions[level]."""
    doc = Document(str(path))
    text = "\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())
    text = text.replace("[Done]", "")

    # Split into two sections if headings exist.
    inter_part = text
    intro_part = text
    m_intro = re.search(r"\bIntroduction\s*:", text, flags=re.I)
    if m_intro:
        inter_part = text[:m_intro.start()]
        intro_part = text[m_intro.end():]
    m_inter = re.search(r"\bWork\s*interaction\s*:", inter_part, flags=re.I)
    if m_inter:
        inter_part = inter_part[m_inter.end():]

    def extract_by_level(section_text: str) -> Dict[str, str]:
        result: Dict[str, str] = {}
        # Matches P-2:, P-3:, P-4:, P-5:, D-1:, NO-A: etc.
        label_re = re.compile(r"(?im)(?:^|\n)\s*(P-\d|D-1|NO-[A-D])\s*:\s*")
        matches = list(label_re.finditer(section_text))
        for i, m in enumerate(matches):
            level = _normalize_level(m.group(1))
            start = m.end()
            end = matches[i + 1].start() if i + 1 < len(matches) else len(section_text)
            body = section_text[start:end].strip()
            body = re.sub(r"\s+", " ", body)
            if body:
                result[level] = body
        return result

    interactions = extract_by_level(inter_part)
    introductions = extract_by_level(intro_part)
    return interactions, introductions


# -----------------------------------------------------------------------------
# Parse the level-differentiated Excel workbook
# -----------------------------------------------------------------------------

def _cell_text(ws: Worksheet, row: int, col: int) -> str:
    v = ws.cell(row, col).value
    return "" if v is None else str(v).strip()


def _find_header_rows(ws: Worksheet) -> List[int]:
    rows = []
    for r in range(1, ws.max_row + 1):
        row_values = [_cell_text(ws, r, c).lower().replace("\n", " ") for c in range(1, min(ws.max_column, 12) + 1)]
        if "theme" in row_values and "level" in row_values and any("responsibility" in x for x in row_values):
            rows.append(r)
    return rows


def _header_map(ws: Worksheet, header_row: int) -> Dict[str, int]:
    mapping: Dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        value = _cell_text(ws, header_row, c).lower().replace("\n", " ")
        if value == "theme":
            mapping["theme"] = c
        elif value == "level":
            mapping["level"] = c
        elif "resp" in value and "code" in value:
            mapping["code"] = c
        elif value == "responsibility" or "responsibility" in value:
            mapping["responsibility"] = c
        elif "standard" in value or "mandatory" in value or "grouping" in value:
            mapping["mandatory"] = c
    return mapping


def _detect_specialty_name(ws: Worksheet, header_row: int) -> str:
    # If the table has a title above it, use it as specialty; otherwise General.
    for r in range(header_row - 1, max(0, header_row - 6), -1):
        values = [_cell_text(ws, r, c) for c in range(1, min(ws.max_column, 8) + 1)]
        joined = " ".join(v for v in values if v)
        if joined and not re.search(r"security coordination officer|level differentiated|mapping", joined, re.I):
            return joined[:80]
    return "General"


def _is_marked(value) -> bool:
    if value is None:
        return False
    s = str(value).strip().lower()
    return s in {"x", "yes", "y", "1", "true", "mandatory", "required"}


def _theme_color(theme: str) -> str:
    if not theme:
        return DEFAULT_THEME_COLOR
    clean = re.sub(r"\s+", " ", theme.strip())
    if clean in THEME_COLORS:
        return THEME_COLORS[clean]
    for known, color in THEME_COLORS.items():
        if clean.lower() == known.lower():
            return color
    return DEFAULT_THEME_COLOR


def parse_level_tables_from_excel(path: str | Path) -> Dict[str, List[LevelTables]]:
    """
    Parse the SCO_KSA_Level_Differentiated workbook.

    The parser supports one or multiple tables per sheet. It looks for rows with headers:
    Theme | Level | Resp. Code | Responsibility | OHR Standard | KSA columns...
    """
    wb = load_workbook(path, data_only=False)
    output: Dict[Tuple[str, str], Dict] = {}

    for ws in wb.worksheets:
        header_rows = _find_header_rows(ws)
        if not header_rows:
            continue
        header_rows.append(ws.max_row + 2)

        for idx, header_row in enumerate(header_rows[:-1]):
            next_header = header_rows[idx + 1]
            hmap = _header_map(ws, header_row)
            if not {"theme", "level", "responsibility"}.issubset(hmap):
                continue

            specialty = _detect_specialty_name(ws, header_row)
            ksa_start = max(hmap.values()) + 1
            # In the uploaded file, KSA headers are in row 4 while table headers are row 2.
            ksa_text_row = header_row + 2 if header_row + 2 <= ws.max_row else header_row
            ksa_cols: List[Tuple[int, str]] = []
            for c in range(ksa_start, ws.max_column + 1):
                ksa_text = _cell_text(ws, ksa_text_row, c)
                if ksa_text and len(ksa_text) > 5:
                    ksa_cols.append((c, ksa_text))

            current_theme = ""
            for r in range(header_row + 3, min(next_header, ws.max_row + 1)):
                level = _normalize_level(_cell_text(ws, r, hmap["level"]))
                responsibility_text = _cell_text(ws, r, hmap["responsibility"])
                if not level or not responsibility_text:
                    continue
                raw_theme = _cell_text(ws, r, hmap["theme"])
                if raw_theme:
                    current_theme = raw_theme
                theme = current_theme or "General"
                code = _cell_text(ws, r, hmap.get("code", 0)) if hmap.get("code") else ""
                mandatory = False
                if hmap.get("mandatory"):
                    mandatory = _is_marked(ws.cell(r, hmap["mandatory"]).value)
                # Bold source responsibility is also treated as mandatory.
                try:
                    mandatory = mandatory or bool(ws.cell(r, hmap["responsibility"]).font.bold)
                except Exception:
                    pass

                key = (level, specialty)
                if key not in output:
                    output[key] = {"responsibilities": [], "ksa_map": {}}
                resp_no = len(output[key]["responsibilities"]) + 1
                output[key]["responsibilities"].append(
                    Responsibility(number=resp_no, theme=theme, text=responsibility_text, source_code=code, mandatory=mandatory)
                )
                for c, ksa_text in ksa_cols:
                    if _is_marked(ws.cell(r, c).value):
                        output[key]["ksa_map"].setdefault(ksa_text, []).append(resp_no)

    result: Dict[str, List[LevelTables]] = {}
    for (level, specialty), data in output.items():
        ksa_rows = [
            KsaRow(number=i + 1, text=ksa_text, responsibilities=resp_numbers)
            for i, (ksa_text, resp_numbers) in enumerate(data["ksa_map"].items())
        ]
        result.setdefault(level, []).append(
            LevelTables(level=level, specialty=specialty, responsibilities=data["responsibilities"], ksa_rows=ksa_rows)
        )
    return result


def get_tables_for_level(all_tables: Dict[str, List[LevelTables]], target_level: str) -> List[LevelTables]:
    """Return tables for target level, with NO-level fallback to equivalent P level."""
    direct = all_tables.get(target_level)
    if direct:
        return direct
    equivalent = LEVEL_EQUIVALENT.get(target_level)
    if equivalent and equivalent in all_tables:
        copied: List[LevelTables] = []
        for tbl in all_tables[equivalent]:
            copied.append(LevelTables(level=target_level, specialty=tbl.specialty, responsibilities=tbl.responsibilities, ksa_rows=tbl.ksa_rows))
        return copied
    return []


# -----------------------------------------------------------------------------
# Word table formatting helpers
# -----------------------------------------------------------------------------

def _set_cell_fill(cell, fill_hex: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill_hex.replace("#", ""))


def _set_cell_borders(cell, color="000000", size="6") -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_borders = tc_pr.first_child_found_in("w:tcBorders")
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
        tc_pr.append(tc_borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = "w:" + edge
        element = tc_borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            tc_borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)


def _set_cell_width(cell, width_inches: float) -> None:
    cell.width = Inches(width_inches)
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.find(qn("w:tcW"))
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(int(width_inches * 1440)))
    tc_w.set(qn("w:type"), "dxa")


def _format_cell_text(cell, font_size=8, bold=False, font_color: Optional[str] = None, align: Optional[int] = None) -> None:
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for paragraph in cell.paragraphs:
        if align is not None:
            paragraph.alignment = align
        for run in paragraph.runs:
            run.font.size = Pt(font_size)
            run.bold = bold
            if font_color:
                run.font.color.rgb = _rgb(font_color)


def _rgb(hex_color: str):
    from docx.shared import RGBColor
    h = hex_color.replace("#", "")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _block_element(block):
    """Return the underlying XML element for a Paragraph or Table."""
    if hasattr(block, "_p"):
        return block._p
    if hasattr(block, "_tbl"):
        return block._tbl
    raise TypeError(f"Unsupported block type: {type(block)}")


def _insert_paragraph_after(block, text: str = "", style=None):
    """
    Insert a paragraph immediately after an existing paragraph/table.

    This version is order-safe. The previous implementation always inserted after
    the same anchor, which could reverse the visual order when several blocks were
    inserted one after another.
    """
    parent = block._parent
    new_para = parent.add_paragraph()
    if text:
        new_para.add_run(text)
    if style:
        new_para.style = style

    _block_element(block).addnext(new_para._p)
    return new_para


def _insert_paragraph_before(paragraph, text: str = "", style=None):
    """Insert a paragraph immediately before an existing paragraph."""
    parent = paragraph._parent
    new_para = parent.add_paragraph()
    if text:
        new_para.add_run(text)
    if style:
        new_para.style = style

    paragraph._p.addprevious(new_para._p)
    return new_para


def _insert_table_after(block, rows: int, cols: int):
    """Insert a table immediately after an existing paragraph/table."""
    parent = block._parent
    table = parent.add_table(rows=rows, cols=cols, width=Inches(7.3))
    _block_element(block).addnext(table._tbl)
    return table


def _remove_element(element) -> None:
    element.getparent().remove(element)


def _paragraph_is_heading(paragraph, name: str) -> bool:
    return paragraph.text.strip().lower().startswith(name.lower())


def _clear_tables_after_heading_until(doc: DocumentObject, start_heading: str, stop_headings: Iterable[str]) -> Optional[object]:
    """
    Remove tables after a heading until a stop heading is reached.
    Returns the heading paragraph object.
    """
    body = doc._body._element
    children = list(body)
    start_idx = None
    for i, child in enumerate(children):
        if child.tag.endswith("}p"):
            p = None
            # Map XML paragraph to python-docx paragraph.
            for para in doc.paragraphs:
                if para._p is child:
                    p = para
                    break
            if p is not None and _paragraph_is_heading(p, start_heading):
                start_idx = i
                break
    if start_idx is None:
        return None

    # Find heading paragraph object.
    start_para = next((p for p in doc.paragraphs if p._p is children[start_idx]), None)
    stop_lower = tuple(s.lower() for s in stop_headings)
    for child in list(body)[start_idx + 1:]:
        if child.tag.endswith("}p"):
            p = next((para for para in doc.paragraphs if para._p is child), None)
            if p is not None and any(p.text.strip().lower().startswith(s) for s in stop_lower):
                break
        if child.tag.endswith("}tbl"):
            _remove_element(child)
    return start_para


def _find_heading_paragraph(doc: DocumentObject, heading: str) -> Optional[object]:
    for paragraph in doc.paragraphs:
        if _paragraph_is_heading(paragraph, heading):
            return paragraph
    return None


def _replace_section_paragraph_text(doc: DocumentObject, heading: str, replacement_text: str, stop_headings: Iterable[str]) -> None:
    """
    Replace the text immediately under a section heading while keeping the
    heading in its original template position.

    This keeps the template order:
        Work Interactions -> Introduction -> Responsibilities -> Skills

    It removes old paragraphs, including blank spacer paragraphs, until the next
    stop heading/table, then inserts the new paragraph directly after the heading.
    """
    heading_p = _find_heading_paragraph(doc, heading)
    if heading_p is None or not replacement_text:
        return

    body = doc._body._element
    children = list(body)
    start_idx = children.index(heading_p._p)
    stop_lower = tuple(s.lower() for s in stop_headings)

    to_remove = []

    for child in children[start_idx + 1:]:
        if child.tag.endswith("}p"):
            p = next((para for para in doc.paragraphs if para._p is child), None)
            if p is not None:
                p_text = p.text.strip()
                # Stop at the next major section heading.
                if any(p_text.lower().startswith(s) for s in stop_lower):
                    break
                # Remove both old body text and blank spacing paragraphs.
                to_remove.append(child)

        elif child.tag.endswith("}tbl"):
            # Stop before the next table. Responsibilities/Skills tables are handled separately.
            break

    for child in to_remove:
        _remove_element(child)

    new_p = _insert_paragraph_after(heading_p, replacement_text)
    for run in new_p.runs:
        run.font.size = Pt(11)


# -----------------------------------------------------------------------------
# Build and insert Responsibility/KSA tables
# -----------------------------------------------------------------------------

def add_responsibility_table_after(paragraph, responsibilities: List[Responsibility]) -> object:
    table = _insert_table_after(paragraph, rows=len(responsibilities) + 1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    widths = [1.6, 0.5, 5.2]

    headers = ["Area of Work", "No.", "[Responsibilities]"]
    for c, header in enumerate(headers):
        cell = table.cell(0, c)
        cell.text = header
        _set_cell_fill(cell, HEADER_FILL)
        _set_cell_borders(cell)
        _set_cell_width(cell, widths[c])
        _format_cell_text(cell, font_size=8, bold=True, font_color=HEADER_FONT, align=WD_ALIGN_PARAGRAPH.CENTER)

    for r, resp in enumerate(responsibilities, start=1):
        row_values = [resp.theme, str(resp.number), resp.text]
        fill = _theme_color(resp.theme)
        for c, value in enumerate(row_values):
            cell = table.cell(r, c)
            cell.text = value
            _set_cell_borders(cell)
            _set_cell_width(cell, widths[c])
            if c in (0, 1):
                _set_cell_fill(cell, fill)
            _format_cell_text(
                cell,
                font_size=8,
                bold=bool(resp.mandatory),
                align=WD_ALIGN_PARAGRAPH.CENTER if c in (0, 1) else WD_ALIGN_PARAGRAPH.LEFT,
            )
    return table


def _chunks(items: List[int], size: int) -> List[List[int]]:
    return [items[i:i + size] for i in range(0, len(items), size)]


def add_ksa_table_after(paragraph, ksa_rows: List[KsaRow], responsibilities: List[Responsibility], max_numbers_per_line: int = 10) -> object:
    """
    Build the Skills/KSA mapping table.

    Important:
    - Keep ALL mapped responsibility numbers.
    - If one KSA maps to more than max_numbers_per_line responsibilities,
      add continuation rows instead of deleting the extra numbers.
    - The first row for a KSA shows No. and Skills text.
      Continuation rows keep those two cells blank and continue the remaining numbers.
    """
    resp_by_no = {r.number: r for r in responsibilities}

    # Build expanded rows first so we know the total row count.
    expanded_rows = []
    for ksa in ksa_rows:
        nums = list(ksa.responsibilities or [])
        chunks = _chunks(nums, max_numbers_per_line) if nums else [[]]
        for chunk_index, chunk in enumerate(chunks):
            expanded_rows.append((ksa, chunk, chunk_index))

    # 2 fixed columns + 10 responsibility-number columns.
    total_cols = 2 + max_numbers_per_line
    table = _insert_table_after(paragraph, rows=len(expanded_rows) + 1, cols=total_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    widths = [0.35, 2.6] + [0.46] * max_numbers_per_line

    # Header
    header_cells = ["No.", "Skills"] + ["Responsibilities"] + [""] * (max_numbers_per_line - 1)
    for c, header in enumerate(header_cells):
        cell = table.cell(0, c)
        cell.text = header
        _set_cell_fill(cell, HEADER_FILL)
        _set_cell_borders(cell)
        _set_cell_width(cell, widths[c])
        _format_cell_text(
            cell,
            font_size=8,
            bold=True,
            font_color=HEADER_FONT,
            align=WD_ALIGN_PARAGRAPH.CENTER,
        )

    # Body
    for r, (ksa, chunk, chunk_index) in enumerate(expanded_rows, start=1):
        # Only the first visual row displays KSA number and text.
        table.cell(r, 0).text = str(ksa.number) if chunk_index == 0 else ""
        table.cell(r, 1).text = ksa.text if chunk_index == 0 else ""

        for c in (0, 1):
            _set_cell_borders(table.cell(r, c))
            _set_cell_width(table.cell(r, c), widths[c])
            _format_cell_text(
                table.cell(r, c),
                font_size=7,
                align=WD_ALIGN_PARAGRAPH.CENTER if c == 0 else WD_ALIGN_PARAGRAPH.LEFT,
            )

        # Fill responsibility numbers for this chunk.
        for offset in range(max_numbers_per_line):
            c = 2 + offset
            cell = table.cell(r, c)

            if offset < len(chunk):
                n = chunk[offset]
                cell.text = str(n)
                resp = resp_by_no.get(n)
                _set_cell_fill(cell, _theme_color(resp.theme if resp else ""))
            else:
                cell.text = ""

            _set_cell_borders(cell)
            _set_cell_width(cell, widths[c])
            _format_cell_text(cell, font_size=7, align=WD_ALIGN_PARAGRAPH.CENTER)

    return table


def _is_general_specialty_name(name: str) -> bool:
    return not str(name or "").strip() or str(name).strip().lower() in {"general", "generalist"}


def _format_section_heading(paragraph) -> None:
    if paragraph.runs:
        paragraph.runs[0].bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER


def replace_tables_in_document(doc: DocumentObject, level_tables: List[LevelTables]) -> None:
    """
    Replace the template Responsibilities and Skills tables without changing the
    template section order.

    The generated document follows the template order exactly:

        Education
        Work Experience
        Work Interactions
        Introduction
        Responsibilities
        Skills

    For the first specialty, the existing template Responsibilities and Skills
    headings are reused in place. Extra specialty blocks, if any, are appended
    after the first Skills table in the same internal order:
        Specialty -> Responsibilities -> Skills.
    """
    if not level_tables:
        return

    # Keep the original headings exactly where they are in the template.
    # Only remove the old tables below those headings.
    resp_heading = _clear_tables_after_heading_until(doc, "Responsibilities", ["Skills"])
    skills_heading = _clear_tables_after_heading_until(doc, "Skills", ["**Note", "Note", "Education", "Work Experience", "Work Interactions", "Introduction", "Responsibilities"])

    if resp_heading is None:
        resp_heading = doc.add_paragraph("Responsibilities")
        _format_section_heading(resp_heading)
    if skills_heading is None:
        skills_heading = doc.add_paragraph("Skills")
        _format_section_heading(skills_heading)

    # First specialty uses the original template locations.
    first = level_tables[0]

    # If the first table has a real specialty name, place it immediately before
    # Responsibilities, so the order remains:
    # Introduction -> Specialty -> Responsibilities -> Skills.
    if not _is_general_specialty_name(first.specialty):
        spec_p = _insert_paragraph_before(resp_heading, first.specialty)
        _format_section_heading(spec_p)

    add_responsibility_table_after(resp_heading, first.responsibilities)
    last_block = add_ksa_table_after(skills_heading, first.ksa_rows, first.responsibilities)

    # Additional specialties are appended after the previous Skills table.
    # Each appended block is inserted sequentially after the previous block
    # to avoid reversed order.
    for tbl in level_tables[1:]:
        if not _is_general_specialty_name(tbl.specialty):
            spec_p = _insert_paragraph_after(last_block, tbl.specialty)
            _format_section_heading(spec_p)
            last_block = spec_p

        resp_p = _insert_paragraph_after(last_block, "Responsibilities")
        _format_section_heading(resp_p)
        last_block = add_responsibility_table_after(resp_p, tbl.responsibilities)

        skill_p = _insert_paragraph_after(last_block, "Skills")
        _format_section_heading(skill_p)
        last_block = add_ksa_table_after(skill_p, tbl.ksa_rows, tbl.responsibilities)


# -----------------------------------------------------------------------------
# Main generation workflow
# -----------------------------------------------------------------------------

def generate_one_document(
    template_path: str | Path,
    target_level: str,
    level_tables: List[LevelTables],
    interactions: Optional[Dict[str, str]] = None,
    introductions: Optional[Dict[str, str]] = None,
) -> bytes:
    doc = Document(str(template_path))
    source_level = detect_template_level(doc)

    replace_level_everywhere(doc, source_level, target_level)
    replace_work_experience_years(doc, target_level)

    # Use equivalent P-level text for NO-levels if exact NO text is not provided.
    lookup_level = target_level
    equivalent = LEVEL_EQUIVALENT.get(target_level)
    interaction_text = (interactions or {}).get(lookup_level) or ((interactions or {}).get(equivalent) if equivalent else None)
    introduction_text = (introductions or {}).get(lookup_level) or ((introductions or {}).get(equivalent) if equivalent else None)

    if interaction_text:
        _replace_section_paragraph_text(doc, "Work Interactions", interaction_text, ["Introduction", "Responsibilities", "Skills"])
    if introduction_text:
        _replace_section_paragraph_text(doc, "Introduction", introduction_text, ["Responsibilities", "Skills"])

    replace_tables_in_document(doc, level_tables)

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def generate_zip(
    excel_path: str | Path,
    template_path: str | Path,
    intro_interaction_path: Optional[str | Path] = None,
) -> Tuple[bytes, List[str]]:
    all_tables = parse_level_tables_from_excel(excel_path)
    interactions: Dict[str, str] = {}
    introductions: Dict[str, str] = {}
    if intro_interaction_path:
        interactions, introductions = parse_intro_interaction_docx(intro_interaction_path)

    warnings: List[str] = []
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for level in OUTPUT_LEVELS:
            tables = get_tables_for_level(all_tables, level)
            if not tables:
                warnings.append(f"No Responsibility/KSA rows found for {level}; generated document keeps template tables.")
            doc_bytes = generate_one_document(
                template_path=template_path,
                target_level=level,
                level_tables=tables,
                interactions=interactions,
                introductions=introductions,
            )
            safe_name = f"Security Coordination Officer GJP 2.0 - {level}.docx"
            zf.writestr(safe_name, doc_bytes)
    return zip_buffer.getvalue(), warnings


# -----------------------------------------------------------------------------
# Streamlit UI
# -----------------------------------------------------------------------------

def _save_uploaded_file(uploaded_file, suffix: str) -> Path:
    tmp_dir = Path(tempfile.mkdtemp(prefix="gjp_doc_generator_"))
    path = tmp_dir / f"uploaded{suffix}"
    path.write_bytes(uploaded_file.getbuffer())
    return path


def render_gjp_document_generator() -> None:
    if st is None:
        raise RuntimeError("Streamlit is not installed. Run: pip install streamlit")
    st.header("Generate GJP Word Documents")
    st.write(
        "Upload the SCO level-differentiated Excel file and a GJP Word template. "
        "The tool will generate P-1, P-2, P-3, P-4, P-5, D-1, NO-A, NO-B, NO-C, and NO-D documents as a ZIP."
    )

    excel_file = st.file_uploader(
        "1) Upload SCO_KSA_Level_Differentiated.xlsx",
        type=["xlsx"],
        key="sco_level_excel",
    )
    template_file = st.file_uploader(
        "2) Upload GJP Word template (.docx)",
        type=["docx"],
        key="gjp_template_docx",
    )
    intro_file = st.file_uploader(
        "3) Optional: upload Work interaction and Introduction Word file (.docx)",
        type=["docx"],
        key="intro_interaction_docx",
    )

    if st.button("Generate GJP Documents", type="primary"):
        if excel_file is None or template_file is None:
            st.error("Please upload both the SCO Excel file and the GJP Word template first.")
            return
        try:
            with st.spinner("Generating Word documents..."):
                excel_path = _save_uploaded_file(excel_file, ".xlsx")
                template_path = _save_uploaded_file(template_file, ".docx")
                intro_path = _save_uploaded_file(intro_file, ".docx") if intro_file is not None else None
                zip_bytes, warnings = generate_zip(excel_path, template_path, intro_path)

            if warnings:
                with st.expander("Generation warnings"):
                    for warning in warnings:
                        st.warning(warning)
            st.success("Done! Download the ZIP file below.")
            st.download_button(
                label="Download generated GJP documents (.zip)",
                data=zip_bytes,
                file_name="generated_gjp_documents.zip",
                mime="application/zip",
            )
        except Exception as exc:
            st.exception(exc)


if __name__ == "__main__":
    render_gjp_document_generator()
