"""
DDL Automation Script
======================
Takes the raw DDL export (as downloaded every 15 days) and produces a
print-ready workbook:
  - "Report" sheet: all consumers, only the columns you care about, in your order
  - One sheet per staff member: same columns, filtered to that staff's DTC codes
  - Landscape A4, fit-to-one-page-wide, Arial 10, bordered header

Usage:
    python process_ddl.py <raw_ddl_file.xls/.xlsx> <staff_mapping.csv> <output.xlsx>

staff_mapping.csv format (you maintain this once, reuse every 15 days):
    dtc_code,staff_name
    4386975,Funde
    4386973,Masaram
    3861624,Masaram
"""

import sys
import subprocess
import os
import csv
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# ---------------------------------------------------------------------------
# CONFIG — edit this section to change which columns appear, and their order.
# Names must match the raw file's header text exactly (trailing spaces included
# is fine, we strip them automatically).
# ---------------------------------------------------------------------------
WANTED_COLUMNS = [
    "Sr No",
    "DTC Code",
    "DTC Section Name",
    "Consumer No",
    "Name",
    "Address",
    "Meter Number ",
    "Bill Due Date",
    "Last Receipt Date",
    "Bill Amount",
    "Closing Balance",
    "Age in Days",
    "Consumers Mobile Number",
    "Remark",
]

HEADER_ROW_IN_RAW = 7        # row where "Sr No, Circle, ..." headers live
DATA_STARTS_AT_ROW = 8       # first data row in the raw file
INFO_ROWS = 5                # number of banner rows (company name, filters, etc.) to copy as-is

COL_WIDTHS = {  # tuned to match a readable A4 landscape printout
    "Sr No": 6,
    "DTC Code": 12,
    "DTC Section Name": 24,
    "Consumer No": 16,
    "Name": 32,
    "Address": 40,
    "Meter Number ":14,
    "Bill Due Date": 12,
    "Last Receipt Date": 12,
    "Bill Amount": 11,
    "Closing Balance": 12,
    "Age in Days": 9,
    "Consumers Mobile Number": 14,
    "Remark": 30,
}

THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FONT = Font(name="Arial", size=10, bold=True)
DATA_FONT = Font(name="Arial", size=10)
BANNER_FONT = Font(name="Arial", size=10, bold=True)


def ensure_xlsx(path: str) -> str:
    """Old .xls files need converting to .xlsx before openpyxl can read them."""
    if path.lower().endswith(".xlsx"):
        return path
    out_dir = os.getcwd()
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "xlsx", "--outdir", out_dir, path],
        check=True, capture_output=True,
    )
    converted = os.path.join(out_dir, os.path.splitext(os.path.basename(path))[0] + ".xlsx")
    return converted


def load_raw(path: str):
    """Returns (banner_rows, header_list, data_rows) from the raw DDL export."""
    xlsx_path = ensure_xlsx(path)
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    banner_rows = []
    for r in range(1, INFO_ROWS + 1):
        banner_rows.append(ws.cell(row=r, column=1).value)

    raw_headers = [
        (ws.cell(row=HEADER_ROW_IN_RAW, column=c).value or "").strip()
        for c in range(1, ws.max_column + 1)
    ]

    col_index = {}
    for name in WANTED_COLUMNS:
        try:
            col_index[name] = raw_headers.index(name) + 1  # 1-based
        except ValueError:
            raise ValueError(
                f"Column '{name}' not found in raw file headers: {raw_headers}"
            )

    data_rows = []
    for r in range(DATA_STARTS_AT_ROW, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=col_index[name]).value for name in WANTED_COLUMNS]
        if all(v is None for v in row_vals):
            continue
        data_rows.append(row_vals)

    return banner_rows, WANTED_COLUMNS, data_rows


def load_staff_map(path: str):
    """dtc_code -> staff_name, from a 2-column CSV."""
    mapping = {}
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            dtc = str(row["dtc_code"]).strip()
            staff = row["staff_name"].strip()
            mapping[dtc] = staff
    return mapping


def style_sheet(ws: Worksheet, headers, n_data_rows: int):
    n_cols = len(headers)

    # Banner rows (1..INFO_ROWS) merged across all columns
    for r in range(1, INFO_ROWS + 1):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=n_cols)
        cell = ws.cell(row=r, column=1)
        cell.font = BANNER_FONT
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    header_row = INFO_ROWS + 2  # one blank row after banner
    for c, name in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c, value=name)
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        letter = get_column_letter(c)
        ws.column_dimensions[letter].width = COL_WIDTHS.get(name, 14)

    for r in range(header_row + 1, header_row + 1 + n_data_rows):
        for c in range(1, n_cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = DATA_FONT
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.border = BORDER

    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)

    # Print setup: landscape, A4, everything on one page width
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5
    ws.print_options.horizontalCentered = True


def write_sheet(wb, sheet_name, banner_rows, headers, rows):
    ws = wb.create_sheet(title=sheet_name[:31])  # Excel sheet name limit
    for r, val in enumerate(banner_rows, start=1):
        ws.cell(row=r, column=1, value=val)

    header_row = INFO_ROWS + 2
    for c, name in enumerate(headers, start=1):
        ws.cell(row=header_row, column=c, value=name)

    for i, row_vals in enumerate(rows):
        for c, val in enumerate(row_vals, start=1):
            ws.cell(row=header_row + 1 + i, column=c, value=val)

    style_sheet(ws, headers, len(rows))
    return ws


def main(raw_path, staff_map_path, out_path):
    banner_rows, headers, data_rows = load_raw(raw_path)
    staff_map = load_staff_map(staff_map_path) if staff_map_path else {}

    dtc_idx = headers.index("DTC Code")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # drop default blank sheet

    # 1) Master "Report" sheet — everyone
    write_sheet(wb, "Report", banner_rows, headers, data_rows)

    # 2) One sheet per staff member
    by_staff = {}
    for row in data_rows:
        dtc = str(row[dtc_idx]).strip()
        staff = staff_map.get(dtc)
        if staff:
            by_staff.setdefault(staff, []).append(row)

    for staff, rows in by_staff.items():
        write_sheet(wb, staff, banner_rows, headers, rows)

    wb.save(out_path)
    print(f"Saved: {out_path}")
    print(f"Report sheet: {len(data_rows)} rows")
    for staff, rows in by_staff.items():
        print(f"  {staff}: {len(rows)} rows")


if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python process_ddl.py <raw_ddl_file> <staff_mapping.csv> <output.xlsx>")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2], sys.argv[3])
