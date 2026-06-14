"""
GJP document generator page for Streamlit.

What it does
------------
1) Accepts either one KSA Excel file or a ZIP containing specialty Excel files.
2) Uploads one GJP Word template.
3) Uploads a Word file containing Work interaction and Introduction text.
4) Generates only levels that contain both Output_<level>_Resp and Output_<level>_KSA data.
5) Replaces level text, updates work experience years, replaces Work Interactions and Introduction,
   and inserts specialty tables in the same order as the Word template.

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

# NO levels may reuse P-level narrative text, but never Responsibility/KSA tables.
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
    fill_hex: str = ""


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
    header_title: str = ""


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
    """Return only exact uploaded data; never manufacture a missing level."""
    return all_tables.get(target_level, [])


OUTPUT_SHEET_RE = re.compile(
    r"^output[\s_-]*(p[\s_-]?[1-5]|d[\s_-]?1|no[\s_-]?[a-d]).*?[\s_-](resp|ksa)$",
    re.I,
)


def _cell_fill_hex(cell) -> str:
    fill = cell.fill
    if not fill or fill.fill_type != "solid":
        return ""
    color = fill.fgColor
    value = color.rgb if color.type == "rgb" else color.indexed
    if value is None:
        return ""
    value = str(value)
    return value[-6:] if len(value) >= 6 else ""


def _specialty_from_header(header: str, fallback: str) -> str:
    fallback_lower = re.sub(r"\.xlsx$", "", Path(str(fallback)).name, flags=re.I).lower()
    filename_specialties = (
        ("generalist", "Generalist"),
        ("civil affairs", "Civil Affairs"),
        ("info analyst", "Information Analyst"),
        ("information analyst", "Information Analyst"),
        ("joint ops", "Joint Operations"),
        ("joint operations", "Joint Operations"),
        ("special assistant", "Special Assistant"),
        ("coordination", "Coordination"),
    )
    for marker, specialty in filename_specialties:
        if marker in fallback_lower:
            return specialty

    text = re.sub(r"\b(P\s*-?\s*[1-5]|D\s*-?\s*1|NO\s*-?\s*[A-D])\b", "", header, flags=re.I)
    text = re.sub(r"\bResponsibilities\b", "", text, flags=re.I)
    text = re.sub(r"\b(Associate|Senior|Assistant|Chief of Service)\b", "", text, flags=re.I)
    text = re.sub(r"\bOfficer\b", "", text, flags=re.I)
    text = re.sub(r"\s+", " ", text).strip(" -_:")
    if text:
        return text

    fallback = re.sub(r"(?i)\b(transposing|ksa|mapping|theme|version|copy|output|6\.2|v\d+)\b", " ", fallback)
    fallback = re.sub(r"[_-]+", " ", fallback)
    return re.sub(r"\s+", " ", fallback).strip() or "General"


def _find_row_with_labels(ws: Worksheet, required: Iterable[str]) -> Optional[int]:
    labels = tuple(x.lower() for x in required)
    for row in range(1, min(ws.max_row, 30) + 1):
        values = [
            _cell_text(ws, row, col).lower().replace("\n", " ")
            for col in range(1, min(ws.max_column, 12) + 1)
        ]
        if all(any(label in value for value in values) for label in labels):
            return row
    return None


def _is_usable_output_text(value: str) -> bool:
    text = str(value or "").strip()
    if not text or text.lower() in {"0", "[]", "[responsibilities]", "[ksa]"}:
        return False
    return not text.startswith("#")


def _parse_output_responsibilities(ws: Worksheet, source_name: str) -> Tuple[str, str, List[Responsibility]]:
    header_row = _find_row_with_labels(ws, ("area of work", "responsib"))
    if header_row is None:
        return "", "", []

    area_col = number_col = text_col = None
    for col in range(1, ws.max_column + 1):
        value = _cell_text(ws, header_row, col).lower()
        if "area of work" in value:
            area_col = col
        elif value in {"no.", "no", "nr", "nr2"}:
            number_col = col
        elif "responsib" in value:
            text_col = col
    if not all((area_col, number_col, text_col)):
        return "", "", []

    header = _cell_text(ws, header_row, text_col)
    specialty = _specialty_from_header(header, source_name)
    responsibilities: List[Responsibility] = []
    blank_run = 0
    current_area = ""
    for row in range(header_row + 1, ws.max_row + 1):
        text = _cell_text(ws, row, text_col)
        if not _is_usable_output_text(text):
            blank_run += 1
            if blank_run >= 4:
                break
            continue
        blank_run = 0
        area = _cell_text(ws, row, area_col)
        if area:
            current_area = area
        raw_number = ws.cell(row, number_col).value
        try:
            number = int(float(raw_number))
        except (TypeError, ValueError):
            number = len(responsibilities) + 1
        text_cell = ws.cell(row, text_col)
        area_cell = ws.cell(row, area_col)
        responsibilities.append(
            Responsibility(
                number=number,
                theme=current_area or "General",
                text=text,
                mandatory=bool(text_cell.font.bold),
                fill_hex=_cell_fill_hex(area_cell),
            )
        )
    return specialty, header, responsibilities


def _parse_output_ksa(ws: Worksheet) -> List[KsaRow]:
    header_row = _find_row_with_labels(ws, ("knowledge, skills", "responsib"))
    if header_row is None:
        return []

    number_col = text_col = mapping_start = None
    for col in range(1, ws.max_column + 1):
        value = _cell_text(ws, header_row, col).lower()
        if value in {"no.", "no", "nr", "nr2"}:
            number_col = col
        elif "knowledge, skills" in value:
            text_col = col
        elif "responsib" in value and mapping_start is None:
            mapping_start = col
    if not all((number_col, text_col, mapping_start)):
        return []

    rows: List[KsaRow] = []
    blank_run = 0
    for row in range(header_row + 1, ws.max_row + 1):
        text = _cell_text(ws, row, text_col)
        if not _is_usable_output_text(text):
            blank_run += 1
            if blank_run >= 4:
                break
            continue
        blank_run = 0
        raw_number = ws.cell(row, number_col).value
        try:
            number = int(float(raw_number))
        except (TypeError, ValueError):
            number = len(rows) + 1
        mapped: List[int] = []
        for col in range(mapping_start, ws.max_column + 1):
            value = ws.cell(row, col).value
            try:
                mapped_number = int(float(value))
            except (TypeError, ValueError):
                continue
            if mapped_number > 0 and mapped_number not in mapped:
                mapped.append(mapped_number)
        rows.append(KsaRow(number=number, text=text, responsibilities=mapped))
    return rows


def parse_output_workbook(path: str | Path) -> Dict[str, List[LevelTables]]:
    """Read the cached values from Output_<level>_Resp and Output_<level>_KSA sheets."""
    wb = load_workbook(path, data_only=True)
    pairs: Dict[str, Dict[str, Worksheet]] = {}
    for ws in wb.worksheets:
        match = OUTPUT_SHEET_RE.match(ws.title.strip())
        if not match:
            continue
        level = _normalize_level(match.group(1))
        kind = match.group(2).lower()
        pairs.setdefault(level, {})[kind] = ws

    result: Dict[str, List[LevelTables]] = {}
    for level, sheets in pairs.items():
        if "resp" not in sheets or "ksa" not in sheets:
            continue
        specialty, header_title, responsibilities = _parse_output_responsibilities(sheets["resp"], str(path))
        ksa_rows = _parse_output_ksa(sheets["ksa"])
        if responsibilities and ksa_rows:
            result.setdefault(level, []).append(
                LevelTables(
                    level=level,
                    specialty=specialty,
                    responsibilities=responsibilities,
                    ksa_rows=ksa_rows,
                    header_title=header_title,
                )
            )
    return result


def parse_level_tables_from_upload(path: str | Path) -> Tuple[Dict[str, List[LevelTables]], List[str]]:
    """Accept one XLSX or a ZIP of XLSX files and combine all available specialties."""
    source = Path(path)
    workbook_paths: List[Path] = []
    temp_dir: Optional[Path] = None
    if source.suffix.lower() == ".zip":
        temp_dir = Path(tempfile.mkdtemp(prefix="gjp_excel_zip_"))
        with zipfile.ZipFile(source) as archive:
            for member in archive.infolist():
                member_path = Path(member.filename)
                if member.is_dir() or member_path.suffix.lower() != ".xlsx" or "__MACOSX" in member_path.parts:
                    continue
                safe_name = f"{len(workbook_paths):03d}_{member_path.name}"
                extracted = temp_dir / safe_name
                extracted.write_bytes(archive.read(member))
                workbook_paths.append(extracted)
    else:
        workbook_paths = [source]

    if not workbook_paths:
        raise ValueError("No .xlsx files were found in the uploaded file.")

    combined: Dict[str, List[LevelTables]] = {}
    warnings: List[str] = []
    for workbook_path in workbook_paths:
        parsed = parse_output_workbook(workbook_path)
        if not parsed:
            warnings.append(f"No usable Output_<level>_Resp/KSA sheet pairs found in {workbook_path.name}.")
            continue
        for level, tables in parsed.items():
            combined.setdefault(level, []).extend(tables)
    return combined, warnings


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

def add_responsibility_table_after(
    paragraph,
    responsibilities: List[Responsibility],
    header_title: str = "[Responsibilities]",
) -> object:
    table = _insert_table_after(paragraph, rows=len(responsibilities) + 1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    widths = [1.6, 0.5, 5.2]

    headers = ["Area of Work", "No.", header_title]
    for c, header in enumerate(headers):
        cell = table.cell(0, c)
        cell.text = header
        _set_cell_fill(cell, HEADER_FILL)
        _set_cell_borders(cell)
        _set_cell_width(cell, widths[c])
        _format_cell_text(cell, font_size=8, bold=True, font_color=HEADER_FONT, align=WD_ALIGN_PARAGRAPH.CENTER)

    for r, resp in enumerate(responsibilities, start=1):
        row_values = [resp.theme, str(resp.number), resp.text]
        fill = resp.fill_hex or _theme_color(resp.theme)
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


def add_ksa_table_after(
    paragraph,
    ksa_rows: List[KsaRow],
    responsibilities: List[Responsibility],
    header_title: str = "Responsibilities",
    max_numbers_per_line: Optional[int] = None,
) -> object:
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
    if max_numbers_per_line is None:
        max_numbers_per_line = max(
            1,
            max((len(row.responsibilities) for row in ksa_rows), default=1),
        )

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

    number_width = max(0.22, 4.35 / max_numbers_per_line)
    widths = [0.35, 2.6] + [number_width] * max_numbers_per_line

    # Header
    header_cells = ["No.", "Knowledge, Skills, and Abilities", header_title] + [""] * (max_numbers_per_line - 1)
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
            mapped_resp = None

            if offset < len(chunk):
                n = chunk[offset]
                cell.text = str(n)
                mapped_resp = resp_by_no.get(n)
                _set_cell_fill(
                    cell,
                    (mapped_resp.fill_hex if mapped_resp else "")
                    or _theme_color(mapped_resp.theme if mapped_resp else ""),
                )
            else:
                cell.text = ""

            _set_cell_borders(cell)
            _set_cell_width(cell, widths[c])
            _format_cell_text(
                cell,
                font_size=7,
                bold=bool(mapped_resp and mapped_resp.mandatory),
                align=WD_ALIGN_PARAGRAPH.CENTER,
            )

    if max_numbers_per_line > 1:
        merged = table.cell(0, 2).merge(table.cell(0, total_cols - 1))
        merged.text = header_title
        _set_cell_fill(merged, HEADER_FILL)
        _set_cell_borders(merged)
        _format_cell_text(
            merged,
            font_size=8,
            bold=True,
            font_color=HEADER_FONT,
            align=WD_ALIGN_PARAGRAPH.CENTER,
        )

    return table


def _is_general_specialty_name(name: str) -> bool:
    return not str(name or "").strip() or str(name).strip().lower() in {"general", "generalist"}


def _format_section_heading(paragraph) -> None:
    if paragraph.runs:
        paragraph.runs[0].bold = True
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _canonical_specialty(value: str) -> str:
    text = str(value or "").lower()
    text = re.sub(r"\bspecialt(?:y|ies)\b", " ", text)
    text = re.sub(r"\b(associate|assistant|senior|chief|service|officer|the)\b", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    aliases = {
        "general": "generalist",
        "generalist": "generalist",
        "info analyst": "information analyst",
        "information analysis": "information analyst",
        "joint ops": "joint operations",
        "coordination": "joint operations",
    }
    return aliases.get(text, text)


def _specialty_match_score(template_name: str, uploaded_name: str) -> int:
    template_key = _canonical_specialty(template_name)
    uploaded_key = _canonical_specialty(uploaded_name)
    if not template_key or not uploaded_key:
        return 0
    if template_key == uploaded_key:
        return 100
    if template_key in uploaded_key or uploaded_key in template_key:
        return 70
    template_words = set(template_key.split())
    uploaded_words = set(uploaded_key.split())
    return len(template_words & uploaded_words) * 10


def _specialty_blocks(doc: DocumentObject) -> List[Tuple[str, object, int, int]]:
    body = doc._body._element
    children = list(body)
    paragraph_by_xml = {p._p: p for p in doc.paragraphs}
    starts: List[Tuple[str, object, int]] = []
    for index, child in enumerate(children):
        paragraph = paragraph_by_xml.get(child)
        if paragraph is None:
            continue
        match = re.match(r"\s*Specialt(?:y|ies)\s*:\s*(.+)", paragraph.text, flags=re.I)
        if match:
            starts.append((match.group(1).strip(), paragraph, index))
    blocks: List[Tuple[str, object, int, int]] = []
    for index, (name, paragraph, start) in enumerate(starts):
        end = starts[index + 1][2] if index + 1 < len(starts) else len(children)
        blocks.append((name, paragraph, start, end))
    return blocks


def _replace_specialty_blocks(
    doc: DocumentObject,
    level_tables: List[LevelTables],
    target_level: str,
    gjp_area: str,
) -> None:
    blocks = _specialty_blocks(doc)
    if not blocks:
        replace_tables_in_document(doc, level_tables, target_level, gjp_area)
        return

    remaining = list(level_tables)
    assignments: List[Tuple[Tuple[str, object, int, int], Optional[LevelTables]]] = []
    for block in blocks:
        best = None
        best_score = 0
        for table_data in remaining:
            score = _specialty_match_score(block[0], table_data.specialty)
            if score > best_score:
                best = table_data
                best_score = score
        if best is not None and best_score >= 10:
            remaining.remove(best)
            assignments.append((block, best))
        else:
            assignments.append((block, None))

    if len(blocks) == 1 and len(level_tables) == 1 and assignments[0][1] is None:
        assignments[0] = (blocks[0], level_tables[0])

    body = doc._body._element
    original_children = list(body)
    paragraph_by_xml = {p._p: p for p in doc.paragraphs}

    # Delete template specialty blocks that have no uploaded data for this level.
    for (name, paragraph, start, end), table_data in reversed(assignments):
        if table_data is None:
            for child in original_children[start:end]:
                if child.getparent() is body:
                    _remove_element(child)

    # Replace the two tables inside each retained specialty block.
    for (name, specialty_paragraph, start, end), table_data in assignments:
        if table_data is None or specialty_paragraph._p.getparent() is not body:
            continue
        segment = [child for child in original_children[start:end] if child.getparent() is body]
        resp_heading = None
        ksa_heading = None
        for child in segment:
            paragraph = paragraph_by_xml.get(child)
            if paragraph is None:
                continue
            text = paragraph.text.strip().lower()
            if text == "responsibilities":
                resp_heading = paragraph
            elif text == "skills" or text.startswith("knowledge, skills"):
                ksa_heading = paragraph
        for child in segment:
            if child.tag.endswith("}tbl"):
                _remove_element(child)
        if resp_heading is None or ksa_heading is None:
            continue

        header_title = table_data.header_title or f"{target_level} {gjp_area} Responsibilities"
        source_level = detect_template_level(doc)
        header_title = re.sub(
            rf"\b{re.escape(source_level)}\b",
            target_level,
            header_title,
            flags=re.I,
        )
        add_responsibility_table_after(resp_heading, table_data.responsibilities, header_title)
        add_ksa_table_after(ksa_heading, table_data.ksa_rows, table_data.responsibilities, header_title)


def replace_tables_in_document(
    doc: DocumentObject,
    level_tables: List[LevelTables],
    target_level: str = "",
    gjp_area: str = "",
) -> None:
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

    first_header = first.header_title or f"{target_level} {gjp_area} Responsibilities".strip()
    add_responsibility_table_after(resp_heading, first.responsibilities, first_header)
    last_block = add_ksa_table_after(skills_heading, first.ksa_rows, first.responsibilities, first_header)

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
        table_header = tbl.header_title or f"{target_level} {gjp_area} Responsibilities".strip()
        last_block = add_responsibility_table_after(resp_p, tbl.responsibilities, table_header)

        skill_p = _insert_paragraph_after(last_block, "Skills")
        _format_section_heading(skill_p)
        last_block = add_ksa_table_after(skill_p, tbl.ksa_rows, tbl.responsibilities, table_header)


# -----------------------------------------------------------------------------
# Main generation workflow
# -----------------------------------------------------------------------------

def generate_one_document(
    template_path: str | Path,
    target_level: str,
    level_tables: List[LevelTables],
    gjp_area: str,
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
        _replace_section_paragraph_text(
            doc,
            "Work Interactions",
            interaction_text,
            ["Education", "Work Experience", "Specialty", "Introduction", "Responsibilities", "Skills"],
        )
    if introduction_text:
        _replace_section_paragraph_text(doc, "Introduction", introduction_text, ["Responsibilities", "Skills"])

    _replace_specialty_blocks(doc, level_tables, target_level, gjp_area)

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def _safe_filename_part(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*]+', " ", str(value))
    return re.sub(r"\s+", " ", value).strip(" .") or "GJP"


def generate_documents(
    excel_or_zip_path: str | Path,
    template_path: str | Path,
    intro_interaction_path: str | Path,
    gjp_area: str,
) -> Tuple[Dict[str, bytes], List[str]]:
    all_tables, warnings = parse_level_tables_from_upload(excel_or_zip_path)
    interactions, introductions = parse_intro_interaction_docx(intro_interaction_path)

    documents: Dict[str, bytes] = {}
    area_name = _safe_filename_part(gjp_area)
    for level in OUTPUT_LEVELS:
        tables = all_tables.get(level, [])
        if not tables:
            continue
        doc_bytes = generate_one_document(
            template_path=template_path,
            target_level=level,
            level_tables=tables,
            gjp_area=gjp_area,
            interactions=interactions,
            introductions=introductions,
        )
        documents[f"{area_name} GJP 2.0 - {level}.docx"] = doc_bytes

    if not documents:
        raise ValueError(
            "No documents were generated. Check that the workbook contains populated "
            "Output_<level>_Resp and Output_<level>_KSA sheet pairs."
        )
    return documents, warnings


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
        "Upload one XLSX file, or a ZIP containing the XLSX files for all specialties. "
        "Only levels with populated Responsibility and KSA output sheets will be generated."
    )

    gjp_area = st.text_input(
        "GJP Area / workstream",
        placeholder="e.g. Political Affairs Officer",
        key="gjp_area_workstream",
    )
    excel_file = st.file_uploader(
        "1) Upload specialty Excel ZIP or one Excel file (.zip or .xlsx)",
        type=["zip", "xlsx"],
        key="gjp_level_excel_or_zip",
    )
    template_file = st.file_uploader(
        "2) Upload GJP Word template (.docx)",
        type=["docx"],
        key="gjp_template_docx",
    )
    intro_file = st.file_uploader(
        "3) Upload Work Interactions and Introduction Word file (.docx)",
        type=["docx"],
        key="intro_interaction_docx",
    )

    if st.button("Generate GJP Documents", type="primary"):
        if not gjp_area.strip() or excel_file is None or template_file is None or intro_file is None:
            st.error(
                "Please enter the GJP Area / workstream and upload the Excel/ZIP file, "
                "Word template, and Work Interactions/Introduction Word file."
            )
            return
        try:
            with st.spinner("Generating Word documents..."):
                input_suffix = Path(excel_file.name).suffix.lower()
                excel_path = _save_uploaded_file(excel_file, input_suffix)
                template_path = _save_uploaded_file(template_file, ".docx")
                intro_path = _save_uploaded_file(intro_file, ".docx")
                documents, warnings = generate_documents(
                    excel_path,
                    template_path,
                    intro_path,
                    gjp_area.strip(),
                )

            if warnings:
                with st.expander("Generation warnings"):
                    for warning in warnings:
                        st.warning(warning)
            st.success(f"Done! Generated {len(documents)} Word document(s).")
            for file_name, doc_bytes in documents.items():
                st.download_button(
                    label=f"Download {file_name}",
                    data=doc_bytes,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"download_{file_name}",
                )
        except Exception as exc:
            st.exception(exc)


if __name__ == "__main__":
    render_gjp_document_generator()
