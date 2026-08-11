"""
Usage (CLI):
    python Week _Summary.py Brown_Ale_2026.xlsx --sheets "ก.ค."

    # หลายเดือนรวมกัน (แบบไฟล์ Q2)
    python build_weekday_summary_pandas.py Q2.xlsx --sheets "เม.ย.,พ.ค.,มิ.ย." \
        --profit-cols "24,25,25"   # ระบุ column index ของกำไรสุทธิ แยกตามชีต ถ้าตำแหน่งเลื่อนไม่เท่ากัน
"""

import argparse
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

WEEKDAYS = [
    ("จันทร์", "จ"), ("อังคาร", "อ"), ("พุธ", "พ"), ("พฤหัส", "พฤ"),
    ("ศุกร์", "ศ"), ("เสาร์", "ส"), ("อาทิตย์", "อา"),
]
FONT_NAME = "Arial"


def _load_daily_data(path, sheet, data_start_row, data_end_row,
                      day_col_idx, revenue_col_idx, profit_col_idx):
    """อ่านข้อมูลรายวันจากชีตเดือนหนึ่งชีต คืนค่าเป็น DataFrame: day, revenue, profit"""
    nrows = data_end_row - data_start_row + 1
    df = pd.read_excel(
        path, sheet_name=sheet, header=None,
        skiprows=data_start_row - 1, nrows=nrows,
    )
    out = pd.DataFrame({
        "day": df[day_col_idx],
        "revenue": pd.to_numeric(df[revenue_col_idx], errors="coerce"),  # "-" -> NaN
        "profit": pd.to_numeric(df[profit_col_idx], errors="coerce"),
    })
    return out


def compute_weekday_summary(
    path: str,
    month_sheets: list[str],
    data_start_row: int = 6,
    data_end_row: int = 36,
    day_col_idx: int = 1,       # column B (0-indexed)
    revenue_col_idx: int = 6,   # column G
    profit_col_idx: int = 25,   # column Z
) -> pd.DataFrame:
    """
    รวมข้อมูลจากทุกชีตในเดือน แล้วคำนวณสรุปตามวันในสัปดาห์ (จ.-อา.)
    คืนค่าเป็น DataFrame พร้อมตัวเลขจริง (ไม่ใช่สูตร)
    """
    frames = [
        _load_daily_data(path, s, data_start_row, data_end_row,
                          day_col_idx, revenue_col_idx, profit_col_idx)
        for s in month_sheets
    ]
    all_days = pd.concat(frames, ignore_index=True)

    rows = []
    for th_name, abbr in WEEKDAYS:
        sub = all_days[all_days["day"] == abbr]
        open_days = sub["revenue"].notna().sum()
        closed_days = sub["revenue"].isna().sum()
        total_rev = sub["revenue"].sum()
        total_profit = sub["profit"].sum()
        avg_rev = total_rev / open_days if open_days else 0
        avg_profit = total_profit / open_days if open_days else 0
        rows.append({
            "วัน": th_name, "ตัวย่อ": abbr,
            "จำนวนวันเปิด": int(open_days), "จำนวนวันปิด": int(closed_days),
            "รายรับเฉลี่ย/วัน (฿)": round(avg_rev),
            "กำไรเฉลี่ย/วัน (฿)": round(avg_profit),
            "รายรับรวม (฿)": round(total_rev),
            "กำไรรวม (฿)": round(total_profit),
        })
    return pd.DataFrame(rows)


def write_summary_sheet(source_path: str, output_path: str, summary: pd.DataFrame,
                         month_sheets: list[str], output_sheet: str = "สรุปวันขายดี"):
    """
    อ่านไฟล์ต้นฉบับ (source_path) แบบ read-only แล้วเขียนเฉพาะตาราง (ค่าจริง ไม่ใช่สูตร,
    เริ่มที่ A1, ไม่มีหัวข้อ/คำอธิบาย) ลงเป็นไฟล์ใหม่ (output_path) โดยไม่แตะไฟล์ต้นฉบับเลย
    """
    wb = load_workbook(source_path)   # เปิดมาเพื่อ "อ่าน" เท่านั้น ยังไม่เซฟทับ
    if output_sheet in wb.sheetnames:
        del wb[output_sheet]
    ws = wb.create_sheet(output_sheet, 0)

    header_font = Font(name=FONT_NAME, size=12, bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="2A78D6")
    normal_font = Font(name=FONT_NAME, size=11)
    bold_font = Font(name=FONT_NAME, size=11, bold=True)
    border = Border(*[Side(style="thin", color="CCCCCC")] * 4)
    center = Alignment(horizontal="center", vertical="center")
    right = Alignment(horizontal="right")

    headers = [h.replace(" (฿)", "") for h in summary.columns]
    for i, h in enumerate(headers):
        cell = ws.cell(row=1, column=1 + i, value=h)
        cell.font, cell.fill, cell.alignment, cell.border = header_font, header_fill, center, border

    for r_offset, (_, row) in enumerate(summary.iterrows()):
        r = 2 + r_offset
        for c_offset, col in enumerate(summary.columns):
            c = 1 + c_offset
            val = row[col]
            cell = ws.cell(row=r, column=c, value=val)
            cell.border = border
            if col in ("วัน", "ตัวย่อ", "จำนวนวันเปิด", "จำนวนวันปิด"):
                cell.font = normal_font
                cell.alignment = center
            else:
                cell.font = bold_font if "รวม" in col else normal_font
                cell.alignment = right
                cell.number_format = "#,##0"

    widths = {"A": 12, "B": 10, "C": 14, "D": 14, "E": 18, "F": 18, "G": 16, "H": 16}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    wb.save(output_path)   # เซฟเป็นไฟล์ใหม่ -- ไฟล์ต้นฉบับไม่ถูกแก้เลย


def _cli():
    p = argparse.ArgumentParser(description="Build weekday sales summary (pandas, real values)")
    p.add_argument("path", help="ไฟล์ต้นฉบับ (จะไม่ถูกแก้ไข)")
    p.add_argument("--output", default=None,
                    help="path ไฟล์ใหม่ที่จะสร้าง (default: <ชื่อเดิม>_summary.xlsx)")
    p.add_argument("--sheets", required=True, help="comma-separated month sheet names")
    p.add_argument("--output-sheet", default="สรุปวันขายดี")
    p.add_argument("--data-start-row", type=int, default=6)
    p.add_argument("--data-end-row", type=int, default=36)
    p.add_argument("--day-col-idx", type=int, default=1, help="0-indexed column, B=1")
    p.add_argument("--revenue-col-idx", type=int, default=6, help="0-indexed column, G=6")
    p.add_argument("--profit-col-idx", type=int, default=25, help="0-indexed column, Z=25")
    args = p.parse_args()

    output_path = args.output
    if output_path is None:
        stem, ext = args.path.rsplit(".", 1)
        output_path = f"{stem}_summary.{ext}"

    sheets = [s.strip() for s in args.sheets.split(",")]
    summary = compute_weekday_summary(
        path=args.path, month_sheets=sheets,
        data_start_row=args.data_start_row, data_end_row=args.data_end_row,
        day_col_idx=args.day_col_idx, revenue_col_idx=args.revenue_col_idx,
        profit_col_idx=args.profit_col_idx,
    )
    print(summary.to_string(index=False))
    write_summary_sheet(args.path, output_path, summary, sheets, args.output_sheet)
    print(f"\nDone. ไฟล์ต้นฉบับ '{args.path}' ไม่ถูกแก้ไข -> สร้างไฟล์ใหม่ที่ '{output_path}' แทน (ค่าจริง ไม่ต้อง recalc)")


if __name__ == "__main__":
    _cli()