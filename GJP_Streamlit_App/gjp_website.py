import re
import io
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st
from docx import Document
from openpyxl import load_workbook

# =========================
# Helper functions
# =========================

def clean_text(text):
    """Clean repeated spaces and line breaks."""
    if text is None:
        return ""
    text = str(text).replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def get_docx_text_and_tables(file_bytes):
    """
    Read all paragraphs and table cell texts from a docx file.
    Returns:
        paragraphs: list of paragraph texts
        all_text: combined text
        tables_text: list of rows, each row is list of cell texts
    """
    doc = Document(io.BytesIO(file_bytes))

    paragraphs = []
    for p in doc.paragraphs:
        txt = clean_text(p.text)
        if txt:
            paragraphs.append(txt)

    tables_text = []
    for table in doc.tables:
        for row in table.rows:
            row_values = [clean_text(cell.text) for cell in row.cells]
            if any(row_values):
                tables_text.append(row_values)
                # Also add table cells into paragraph-like text for easier searching
                for cell_text in row_values:
                    if cell_text:
                        paragraphs.append(cell_text)

    all_text = "\n".join(paragraphs)
    return paragraphs, all_text, tables_text


def extract_gjp_job_title(paragraphs):
    """
    Extract GJP Job Title.

    Priority 1:
        From description sentence:
        The Chief of Service, Human Rights D-1 GJP is available for ...
        -> Chief of Service, Human Rights D-1

    Priority 2:
        From line under GENERIC JOB PROFILE:
        GENERIC JOB PROFILE
        Assistant Human Rights Officer: P-1

    Output will also normalize:
        Assistant Human Rights Officer: P-1
        -> Assistant Human Rights Officer P-1
    """

    # Priority 1: extract from "The ... GJP is available for ..."
    for txt in paragraphs:
        text = clean_text(txt)

        match = re.search(
            r"^The\s+(.+?)\s+GJP\s+is\s+available\s+for\b",
            text,
            re.IGNORECASE
        )

        if match:
            title = clean_text(match.group(1))
            title = title.replace(": ", " ")
            title = title.replace(" :", " ")
            title = re.sub(r"\s+", " ", title).strip()
            return title

    # Priority 2: extract the second line under GENERIC JOB PROFILE
    for i, txt in enumerate(paragraphs):
        if clean_text(txt).upper() == "GENERIC JOB PROFILE":
            for j in range(i + 1, min(i + 6, len(paragraphs))):
                candidate = clean_text(paragraphs[j])

                if candidate and not candidate.lower().startswith("version"):
                    candidate = candidate.replace(": ", " ")
                    candidate = candidate.replace(" :", " ")
                    candidate = re.sub(r"\s+", " ", candidate).strip()
                    return candidate

    # Priority 3 fallback: find something like "Officer: P-1" or "Officer P-1"
    pattern = re.compile(r".+\s*:?\s*(P|D|G|NO|FS)-?\d+", re.IGNORECASE)

    for txt in paragraphs:
        text = clean_text(txt)
        if pattern.match(text):
            text = text.replace(": ", " ")
            text = text.replace(" :", " ")
            text = re.sub(r"\s+", " ", text).strip()
            return text

    return ""

def normalize_workstream_name(folder_name):
    """
    Convert GJP Area/ Workstream into lowercase underscore format.
    Example:
        Human Rights -> human_rights
        Political Affairs -> political_affairs
    """
    text = clean_text(folder_name)
    text = text.lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text)
    text = text.strip("_")
    return text


def extract_title_before_level_from_description(paragraphs):
    """
    Extract title before level from:
        The Chief of Service, Human Rights D-1 GJP is available for ...
    Output:
        Chief of Service, Human Rights

    Supported levels:
        P-1, P-2, P-3, P-4, P-5,
        D-1,
        NO-A, NO-B, NO-C, NO-D
    """

    level_pattern = r"(?:P-[1-5]|D-1|NO-[A-D])"

    for txt in paragraphs:
        text = clean_text(txt)

        match = re.search(
            rf"^The\s+(.+?)\s+{level_pattern}\s+GJP\s+is\s+available\s+for\b",
            text,
            re.IGNORECASE
        )

        if match:
            return clean_text(match.group(1))

    return ""


def format_department_name(folder_name):
    """
    Convert folder / department name into required display format.
    Example:
        Human Rights -> Human rights
        Political Affairs -> Political affairs
    """
    text = clean_text(folder_name)
    text = re.sub(r"\s+", " ", text).strip()

    if not text:
        return ""

    # Make first letter uppercase, rest lowercase
    return text[0].upper() + text[1:].lower()


def clean_mapped_specialty_name(specialty, folder_name, paragraphs):
    """
    Rules:
    1. If specialty is General / general / Generalist:
       -> Generalist - [department name]
       Example:
          Human Rights -> Generalist - Human rights

    2. If specialty is blank:
       -> extract title before level from description sentence.
       Example:
          The Chief of Service, Human Rights D-1 GJP is available for
          -> Chief of Service, Human Rights
    """

    specialty = clean_text(specialty)
    specialty = specialty.replace('"', "").replace("“", "").replace("”", "").strip()

    # Rule 1: General / Generalist
    if specialty.lower() in ["general", "generalist"]:
        department_name = format_department_name(folder_name)
        return f"Generalist - {department_name}"

    # Rule 2: blank mapped specialty
    if specialty == "":
        extracted = extract_title_before_level_from_description(paragraphs)
        if extracted:
            return extracted

    return specialty


def extract_job_code(all_text):
    """
    Extract one or multiple Job Codes.
    Examples:
        The CCOG Code is 1.G.02. and the Job Code is 1220.
        The Job Codes are 10225, 3844 and 8624.
    Output:
        1220
        10225, 3844, 8624
    """
    text = clean_text(all_text)

    # Match "Job Code is ..." or "Job Codes are ..."
    match = re.search(
        r"Job Codes?\s+(?:is|are)\s+(.+?)(?:\.|\n|Signed:|The\s)",
        text,
        re.IGNORECASE
    )

    if not match:
        return ""

    segment = match.group(1)

    # Extract all number-like job codes
    codes = re.findall(r"\b\d+\b", segment)

    if codes:
        return ", ".join(codes)

    # fallback
    segment = segment.replace(" and ", ", ")
    segment = re.sub(r"\s+", " ", segment)
    return segment.strip().rstrip(".")


def extract_description(paragraphs):
    """
    Extract the full sentence/paragraph:
    The ... is available for ...
    """
    for txt in paragraphs:
        if re.search(r"^The\s+.+?\s+is available for\s+", txt, re.IGNORECASE):
            return clean_text(txt)

    # fallback: paragraph containing "available for use"
    for txt in paragraphs:
        if "available for use" in txt.lower():
            return clean_text(txt)

    return ""


def extract_section_paragraphs(paragraphs, section_name, stop_sections):
    """
    Extract paragraphs between one heading and the next heading.
    Example:
        section_name = "Education"
        stop_sections = ["Work Experience", "Work Interactions", "Introduction"]
    """
    start_idx = None

    for i, txt in enumerate(paragraphs):
        if clean_text(txt).lower() == section_name.lower():
            start_idx = i
            break

    if start_idx is None:
        return []

    results = []
    stop_lower = [s.lower() for s in stop_sections]

    for txt in paragraphs[start_idx + 1:]:
        current = clean_text(txt)
        if current.lower() in stop_lower:
            break
        if current:
            results.append(current)

    return results


def extract_education_description(paragraphs):
    """
    Education part:
    Usually first paragraph is:
      The education requirement is the same...
    Second paragraph is:
      Advanced university degree ...
    We need the whole second paragraph.
    """
    edu_paras = extract_section_paragraphs(
        paragraphs,
        "Education",
        ["Work Experience", "Work Interactions", "Introduction", "Responsibilities", "Knowledge, Skills, Abilities (KSAs)"]
    )

    if len(edu_paras) >= 2:
        return clean_text(edu_paras[1])
    elif len(edu_paras) == 1:
        return clean_text(edu_paras[0])
    return ""


def extract_minimum_education_level(education_description):
    """
    From:
      Advanced university degree (Master’s degree or equivalent) in human rights, law, ...
    Extract:
      Advanced university degree (Master’s degree or equivalent)
    """
    text = clean_text(education_description)

    # most common split: before " in ..."
    match = re.match(r"(.+?\))\s+in\s+", text, re.IGNORECASE)
    if match:
        return clean_text(match.group(1))

    # fallback: before first " in "
    parts = re.split(r"\s+in\s+", text, maxsplit=1, flags=re.IGNORECASE)
    if parts:
        return clean_text(parts[0])

    return text


def extract_work_experience_description(paragraphs):
    """
    Work Experience part:
    Usually first paragraph is:
      The work experience requirement is the same...
    Second paragraph is the actual requirement.
    We need the whole second paragraph.
    """
    work_paras = extract_section_paragraphs(
        paragraphs,
        "Work Experience",
        ["Work Interactions", "Introduction", "Responsibilities", "Knowledge, Skills, Abilities (KSAs)", "Education"]
    )

    # remove guidance paragraph starting with "In addition..."
    work_paras = [
        p for p in work_paras
        if not p.lower().startswith("in addition")
    ]

    if len(work_paras) >= 2:
        return clean_text(work_paras[1])
    elif len(work_paras) == 1:
        return clean_text(work_paras[0])
    return ""


def extract_required_years(work_experience_description):
    """
    Extract years from:
      A minimum of seven years ...
      A minimum of 7 years ...
    If text says no experience is required, return 0.
    """
    text = clean_text(work_experience_description)

    if re.search(r"not required to have professional work experience|no professional work experience", text, re.IGNORECASE):
        return "0"

    number_words = {
        "zero": "0",
        "one": "1",
        "two": "2",
        "three": "3",
        "four": "4",
        "five": "5",
        "six": "6",
        "seven": "7",
        "eight": "8",
        "nine": "9",
        "ten": "10",
        "eleven": "11",
        "twelve": "12",
        "thirteen": "13",
        "fourteen": "14",
        "fifteen": "15",
    }

    match = re.search(
        r"A minimum of\s+(\d+|zero|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen)\s+years?",
        text,
        re.IGNORECASE
    )

    if match:
        value = match.group(1).lower()
        return number_words.get(value, value)

    return ""


def extract_specialties(all_text):
    """
    Extract mapped specialties from text like:
        Mapped Specialty Name includes the specialties "A", "B", "C"
        includes the specialties "A", "B", "C"
    Remove quotation marks and return list.

    If no specialties are found, return empty list.
    """
    text = clean_text(all_text)

    # capture quoted values after "specialties"
    match = re.search(r"specialties?\s+(.+?)(?:\.|$)", text, re.IGNORECASE)
    if not match:
        return []

    segment = match.group(1)

    quoted = re.findall(r'"([^"]+)"', segment)
    if quoted:
        return [clean_text(x) for x in quoted if clean_text(x)]

    # fallback split by comma / semicolon
    segment = segment.replace('"', "").replace("“", "").replace("”", "")
    items = re.split(r"\s*[,;]\s*", segment)
    return [clean_text(x) for x in items if clean_text(x)]


def extract_one_docx(file_name, file_bytes, folder_name):
    """
    Extract one GJP document into one or multiple rows.
    If specialties exist, put each specialty in a separate row.

    Column names strictly follow the GJP Upload template.
    French fields are temporarily filled with the same English content.
    """
    paragraphs, all_text, tables_text = get_docx_text_and_tables(file_bytes)

    gjp_title = extract_gjp_job_title(paragraphs)
    job_code = extract_job_code(all_text)
    description = extract_description(paragraphs)

    edu_desc = extract_education_description(paragraphs)
    min_edu = extract_minimum_education_level(edu_desc)

    work_desc = extract_work_experience_description(paragraphs)
    required_years = extract_required_years(work_desc)

    specialties = extract_specialties(all_text)

    # If no specialty found, still generate one row with blank Mapped Specialty Name
    if not specialties:
        specialties = [""]

    rows = []

    for specialty in specialties:
        row = {
            "GJP Template ID": "",
            "Status": "Active",
            "GJP Area/ Workstream": folder_name,

            "GJP Job Title": gjp_title,
            "GJP Job Title (French)": gjp_title,

            "Description": description,
            "Description (French)": description,

            "Jobcode ": job_code,

            "Required Years of Work Experience": required_years,

            "Work Experience Description ": work_desc,
            "Work Experience Description (French)": work_desc,

            "Minimum Level of Education Required": min_edu,

            "Education Description (English)": edu_desc,
            "Education Description (French)": edu_desc,

            "Mapped Specialty Name": clean_mapped_specialty_name(specialty=specialty, folder_name=folder_name, paragraphs=paragraphs),
        }

        rows.append(row)

    return rows


def read_uploaded_files(uploaded_files):
    """
    Accept multiple uploaded files.
    Supports:
      - .docx files selected directly
      - .zip file containing docx files
    """
    docx_files = []

    for uploaded in uploaded_files:
        file_name = uploaded.name
        file_bytes = uploaded.read()

        if file_name.lower().endswith(".docx"):
            docx_files.append((file_name, file_bytes))

        elif file_name.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(file_bytes), "r") as z:
                for name in z.namelist():
                    if name.lower().endswith(".docx") and not name.startswith("__MACOSX"):
                        docx_files.append((Path(name).name, z.read(name)))

    return docx_files


def fill_template_if_provided(df, template_file):
    """
    If user uploads an Excel template, fill the columns that match.
    Otherwise, return a normal generated Excel.
    """
    output = io.BytesIO()

    if template_file is None:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="GJP Upload")
        output.seek(0)
        return output

    template_bytes = template_file.read()
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active

    # Find header row by searching for known columns
    known_cols = set(df.columns)
    header_row = None
    header_map = {}

    for row in range(1, min(ws.max_row, 20) + 1):
        temp_map = {}
        for col in range(1, ws.max_column + 1):
            value = ws.cell(row=row, column=col).value
            if value:
                header = clean_text(value)
                temp_map[header] = col

        matched = len(set(temp_map.keys()) & known_cols)
        if matched >= 3:
            header_row = row
            header_map = temp_map
            break

    if header_row is None:
        # fallback: create a new sheet
        ws = wb.create_sheet("Generated GJP Upload")
        for col_idx, col_name in enumerate(df.columns, start=1):
            ws.cell(row=1, column=col_idx).value = col_name
        header_row = 1
        header_map = {col: idx for idx, col in enumerate(df.columns, start=1)}

    start_row = header_row + 1

    # clear existing values under matching columns
    for excel_header, col_idx in header_map.items():
        if excel_header in df.columns:
            for row in range(start_row, ws.max_row + 1):
                ws.cell(row=row, column=col_idx).value = None

    # write data
    for r_idx, record in enumerate(df.to_dict(orient="records"), start=start_row):
        for excel_header, col_idx in header_map.items():
            if excel_header in record:
                ws.cell(row=r_idx, column=col_idx).value = record[excel_header]

    wb.save(output)
    output.seek(0)
    return output


# =========================
# Streamlit website
# =========================

st.set_page_config(page_title="GJP Upload Template Generator", layout="wide")

st.title("GJP Upload Template Generator")
st.write(
    "Upload one department folder as a ZIP file, or select multiple DOCX files. "
    "The app will extract GJP fields and generate an Excel upload template."
)

folder_name = st.text_input(
    "GJP Area / workstream",
    value="Human Rights",
    help="This will be filled into the GJP Area/workstream column. Usually this is the department or folder name."
)

uploaded_files = st.file_uploader(
    "Upload DOCX files or one ZIP folder",
    type=["docx", "zip"],
    accept_multiple_files=True
)

template_file = st.file_uploader(
    "Optional: upload the official GJP Upload template Excel",
    type=["xlsx"]
)

if st.button("Generate Excel"):
    if not uploaded_files:
        st.error("Please upload at least one DOCX file or one ZIP file.")
    else:
        all_rows = []
        docx_files = read_uploaded_files(uploaded_files)

        if not docx_files:
            st.error("No DOCX files found. Please upload DOCX files or a ZIP containing DOCX files.")
        else:
            for file_name, file_bytes in docx_files:
                try:
                    rows = extract_one_docx(file_name, file_bytes, folder_name)
                    all_rows.extend(rows)
                except Exception as e:
                    st.warning(f"Failed to process {file_name}: {e}")

            if not all_rows:
                st.error("No data extracted.")
            else:
                df = pd.DataFrame(all_rows)
                df["GJP Template ID"] = range(1, len(df) + 1)

                st.success(f"Extracted {len(df)} row(s) from {len(docx_files)} document(s).")
                st.dataframe(df, use_container_width=True)

                excel_output = fill_template_if_provided(df, template_file)

                st.download_button(
                    label="Download generated Excel",
                    data=excel_output,
                    file_name="GJP_upload_generated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
