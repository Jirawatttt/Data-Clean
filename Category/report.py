from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import openpyxl
import pandas as pd

from pos_sales_pipeline import (
    load_category_master,
    load_pos_monthly_sales,
    load_receipt_detail_totals,
    load_mapping,
    apply_mapping,
    cross_check_month,
    build_menu_report,
)

MONTH_PATTERN = re.compile(r"_(\d{1,2})_(\d{2})(?:\D|$)")


def extract_month_label(filename: str) -> str:
    """Pull a 'MM-YY' label out of a filename like '..._08_69.xlsx'.
    Falls back to the filename stem if no such pattern is found."""
    m = MONTH_PATTERN.search(filename)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    return Path(filename).stem


def sheet_names(path: Path) -> list[str]:
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
        return wb.sheetnames
    except Exception:
        return []


def classify_file(path: Path):
    """Return one of 'category_master', 'sales', 'receipt', or None."""
    sheets = sheet_names(path)
    if "Category" in sheets:
        return "category_master"
    if "รายละเอียดสินค้าในบิล" in sheets:
        return "receipt"
    # Sales-by-product export: check the header row of the first sheet.
    try:
        df = pd.read_excel(path, sheet_name=0, nrows=0)
        if {"ชื่อสินค้า", "จำนวน"}.issubset(set(df.columns)):
            return "sales"
    except Exception:
        pass
    return None


def discover_files(folder: Path):
    category_master_path = None
    sales_files: dict[str, Path] = {}
    receipt_files: dict[str, Path] = {}

    for path in sorted(folder.glob("*.xlsx")):
        if path.name.startswith("~$"):  # Excel lock files
            continue
        kind = classify_file(path)
        if kind == "category_master":
            if category_master_path is not None:
                print(f"[คำเตือน] เจอไฟล์ master menu มากกว่า 1 ไฟล์ ใช้ไฟล์แรก: {category_master_path.name}")
            else:
                category_master_path = path
        elif kind == "sales":
            label = extract_month_label(path.name)
            sales_files[label] = path
        elif kind == "receipt":
            label = extract_month_label(path.name)
            receipt_files[label] = path

    return category_master_path, sales_files, receipt_files


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("folder", nargs="?", default=".", help="โฟลเดอร์ที่เก็บไฟล์ทั้งหมด (ค่าเริ่มต้น: โฟลเดอร์ปัจจุบัน)")
    parser.add_argument("--category", default=None, help="กรองเฉพาะหมวดหมู่นี้ เช่น Food (ไม่ใส่ = เอาทุกหมวด)")
    parser.add_argument("--mapping-file", default=None, help="path ของ name_mapping.json (ค่าเริ่มต้น: หาในโฟลเดอร์เดียวกัน)")
    args = parser.parse_args()

    folder = Path(args.folder).expanduser().resolve()
    if not folder.is_dir():
        sys.exit(f"ไม่พบโฟลเดอร์: {folder}")

    mapping_file = Path(args.mapping_file) if args.mapping_file else folder / "name_mapping.json"
    if not mapping_file.exists():
        sys.exit(f"ไม่พบไฟล์ name_mapping.json ที่ {mapping_file} (ระบุ --mapping-file ถ้าไฟล์อยู่ที่อื่น)")

    category_master_path, sales_files, receipt_files = discover_files(folder)

    if category_master_path is None:
        sys.exit("ไม่พบไฟล์ master menu (ไฟล์ .xlsx ที่มี sheet ชื่อ 'Category') ในโฟลเดอร์นี้")
    if not sales_files:
        sys.exit("ไม่พบไฟล์ยอดขายตามสินค้าเลยในโฟลเดอร์นี้ (ต้องมีคอลัมน์ ชื่อสินค้า/จำนวน/ยอดขายสุทธิ (฿))")

    print(f"โฟลเดอร์: {folder}")
    print(f"Master menu: {category_master_path.name}")
    print(f"เจอไฟล์ยอดขายตามสินค้า {len(sales_files)} เดือน: {', '.join(sales_files)}")
    if receipt_files:
        print(f"เจอไฟล์ใบเสร็จ {len(receipt_files)} เดือน: {', '.join(receipt_files)} (ใช้ cross-check)")
    else:
        print("ไม่เจอไฟล์ใบเสร็จ -> ข้ามขั้นตอน cross-check")

    category_master = load_category_master(category_master_path)
    mapping = load_mapping(mapping_file)

    monthly_mapped = {}
    all_unmapped = {}
    for month_label, path in sales_files.items():
        raw_sales = load_pos_monthly_sales(path)
        mapped, unmapped = apply_mapping(raw_sales, mapping, category_master)
        if unmapped:
            all_unmapped[month_label] = unmapped
        monthly_mapped[month_label] = mapped.dropna(subset=["master_name"])

        if month_label in receipt_files:
            receipt_totals = load_receipt_detail_totals(receipt_files[month_label])
            mism = cross_check_month(raw_sales, receipt_totals, month_label)
            if not mism.empty:
                print(f"\n[cross-check] เดือน {month_label}: {len(mism)} รายการที่ยอดต่างกันเกิน tolerance:")
                print(mism.to_string(index=False))

    months_sorted = sorted(monthly_mapped)
    cat_suffix = f"_{args.category}" if args.category else ""
    output_path = folder / f"ยอดขายเมนู{cat_suffix}_เดือน{'-'.join(months_sorted)}.xlsx"

    if all_unmapped:
        print("\n=== UNMAPPED PRODUCT NAMES (ต้องเพิ่มใน name_mapping.json) ===")
        for month_label, names in all_unmapped.items():
            print(f"-- เดือน {month_label} --")
            for n in names:
                print(f"   {n}")
        unmapped_flat = pd.DataFrame(
            [(m, n) for m, names in all_unmapped.items() for n in names],
            columns=["month", "pos_name"],
        )
        unmapped_csv = output_path.with_suffix("").with_suffix(".unmapped.csv")
        unmapped_flat.to_csv(unmapped_csv, index=False, encoding="utf-8-sig")
        print(f"\n(รายการนี้ถูกบันทึกไว้ที่ {unmapped_csv.name} ด้วย)")

    report = build_menu_report(category_master, monthly_mapped, category_filter=args.category)
    report.to_excel(output_path, index=False)
    print(f"\n บันทึกรายงานแล้ว: {output_path.name}  ({len(report)} เมนู)")


if __name__ == "__main__":
    main()
