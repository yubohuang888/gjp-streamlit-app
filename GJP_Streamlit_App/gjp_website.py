import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy
import os


# =========================
# 1. 修改这里：输入文件路径
# =========================

input_file = r"C:\Users\YHUANG43\Desktop\SCO_KSA_Level_Differentiated.xlsx"
output_folder = r"C:\Users\YHUANG43\Desktop\GJP_Level_Output"

os.makedirs(output_folder, exist_ok=True)


# =========================
# 2. 工具函数：读取 merged cells 的值
# =========================

def get_merged_cell_value(ws, row, col):
    """
    如果当前 cell 是 merged cell 的一部分，返回 merged range 左上角的值。
    否则返回当前 cell 自己的值。
    """
    cell = ws.cell(row=row, column=col)

    if cell.value is not None:
        return cell.value

    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return ws.cell(
                row=merged_range.min_row,
                column=merged_range.min_col
            ).value

    return None


def clean_text(x):
    if x is None:
        return ""
    return str(x).replace("\n", " ").replace("\r", " ").strip()


# =========================
# 3. 自动找到表头行
# =========================

def find_main_header_row(ws):
    """
    找到包含 Theme / Level / Responsibility 的那一行。
    """
    for row in range(1, min(ws.max_row, 20) + 1):
        values = [
            clean_text(get_merged_cell_value(ws, row, col)).lower()
            for col in range(1, ws.max_column + 1)
        ]

        has_theme = any(v == "theme" for v in values)
        has_level = any(v == "level" for v in values)
        has_responsibility = any(v == "responsibility" for v in values)

        if has_theme and has_level and has_responsibility:
            return row

    raise ValueError(f"Cannot find header row in sheet: {ws.title}")


# =========================
# 4. 从一个 sheet 里筛选某个 level
# =========================

def extract_level_table(input_file, sheet_name, target_level):
    wb = load_workbook(input_file, data_only=True)
    ws = wb[sheet_name]

    main_header_row = find_main_header_row(ws)

    # 第一层表头，比如 Theme, Level, Responsibility, KNOWLEDGE, SKILLS
    main_headers = [
        clean_text(get_merged_cell_value(ws, main_header_row, col))
        for col in range(1, ws.max_column + 1)
    ]

    # K1/K2/K3 那一行
    ksa_code_row = main_header_row + 1

    # 真正的 KSA 文本行，也就是 Knowledge of... / Skills to...
    ksa_text_row = main_header_row + 2

    # 找到基础列位置
    col_map = {}
    for idx, header in enumerate(main_headers, start=1):
        h = clean_text(header).lower()

        if h == "theme":
            col_map["theme"] = idx
        elif h == "level":
            col_map["level"] = idx
        elif "resp" in h and "code" in h:
            col_map["resp_code"] = idx
        elif h == "responsibility":
            col_map["responsibility"] = idx
        elif "ohr" in h and "standard" in h:
            col_map["ohr_standard"] = idx

    required_cols = ["theme", "level", "responsibility"]
    for c in required_cols:
        if c not in col_map:
            raise ValueError(f"Cannot find column '{c}' in sheet {sheet_name}")

    # KSA columns: 从 Responsibility / OHR Standard 后面开始，且有 KSA text 的列
    start_ksa_col = col_map.get("ohr_standard", col_map["responsibility"]) + 1

    ksa_cols = []
    ksa_headers = []

    for col in range(start_ksa_col, ws.max_column + 1):
        ksa_text = clean_text(get_merged_cell_value(ws, ksa_text_row, col))
        ksa_code = clean_text(get_merged_cell_value(ws, ksa_code_row, col))

        # 只保留真正有 KSA 文本的列
        if ksa_text:
            ksa_cols.append(col)
            ksa_headers.append(ksa_text)

    rows = []
    current_theme = ""

    for row in range(ksa_text_row + 1, ws.max_row + 1):
        theme = clean_text(get_merged_cell_value(ws, row, col_map["theme"]))
        level = clean_text(get_merged_cell_value(ws, row, col_map["level"]))
        responsibility = clean_text(get_merged_cell_value(ws, row, col_map["responsibility"]))

        # Theme 可能是合并单元格，也可能只有第一行有值，所以这里 fill down
        if theme:
            current_theme = theme
        else:
            theme = current_theme

        if not responsibility:
            continue

        if level.upper() != target_level.upper():
            continue

        row_data = {
            "Nr2": len(rows) + 1,
            "Grouping": "Optional",
            "Theme": theme,
            "Responsibility": responsibility
        }

        for col, ksa_header in zip(ksa_cols, ksa_headers):
            value = clean_text(get_merged_cell_value(ws, row, col))

            if value.upper() == "X" or value == "×":
                row_data[ksa_header] = "X"
            else:
                row_data[ksa_header] = ""

        rows.append(row_data)

    df = pd.DataFrame(rows)

    return df


# =========================
# 5. 格式化输出 Excel
# =========================

def format_output_excel(file_path, sheet_name="Output"):
    wb = load_workbook(file_path)
    ws = wb[sheet_name]

    # 样式
    header_fill = PatternFill("solid", fgColor="D9D9D9")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # header style
    for cell in ws[1]:
        cell.font = Font(bold=True, size=10)
        cell.fill = header_fill
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )
        cell.border = border

    # body style
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.font = Font(size=10)
            cell.alignment = Alignment(
                horizontal="center" if cell.column not in [3, 4] else "left",
                vertical="center",
                wrap_text=True
            )
            cell.border = border

    # column width
    ws.column_dimensions["A"].width = 6      # Nr2
    ws.column_dimensions["B"].width = 14     # Grouping
    ws.column_dimensions["C"].width = 28     # Theme
    ws.column_dimensions["D"].width = 70     # Responsibility

    # KSA columns
    for col in range(5, ws.max_column + 1):
        col_letter = get_column_letter(col)
        ws.column_dimensions[col_letter].width = 24

    # row height
    ws.row_dimensions[1].height = 85
    for row in range(2, ws.max_row + 1):
        ws.row_dimensions[row].height = 45

    # freeze pane
    ws.freeze_panes = "E2"

    # filter
    ws.auto_filter.ref = ws.dimensions

    wb.save(file_path)


# =========================
# 6. 生成 P2, P4, D1 三个 Excel
# =========================

tasks = [
    {
        "level": "P2",
        "sheet": "P2-P4 Level Differentiated",
        "output": "P2_output.xlsx"
    },
    {
        "level": "P4",
        "sheet": "P2-P4 Level Differentiated",
        "output": "P4_output.xlsx"
    },
    {
        "level": "D1",
        "sheet": "P5-D1 Level Differentiated",
        "output": "D1_output.xlsx"
    }
]

for task in tasks:
    level = task["level"]
    sheet_name = task["sheet"]
    output_name = task["output"]

    print(f"Processing {level} from {sheet_name}...")

    df = extract_level_table(
        input_file=input_file,
        sheet_name=sheet_name,
        target_level=level
    )

    output_path = os.path.join(output_folder, output_name)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Output", index=False)

    format_output_excel(output_path, sheet_name="Output")

    print(f"Saved: {output_path}")
    print(f"Rows exported: {len(df)}")
    print("-" * 50)

print("All files generated successfully!")
