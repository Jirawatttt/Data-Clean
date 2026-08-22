from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import pandas as pd

# --------------------------------------------------------------------------
# 1. LOAD MASTER MENU (Category sheet of Brown_Ale_Cogs.xlsx)
# --------------------------------------------------------------------------

def load_category_master(path: str | Path, sheet_name: str = "Category") -> pd.DataFrame:
    """Parse the Category sheet into a tidy [category, master_name, price] table.

    The sheet is laid out as repeating blocks:
        <Category name>          (row with only column A filled)
        Name | Price | ...       (header row, ignored)
        <item name> | <price>    (one row per menu item)
        ...
        <Category name>          (next block)
        ...
    """
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None)

    records = []
    current_category = None
    for _, row in raw.iterrows():
        col_a, col_b = row[0], row[1] if len(row) > 1 else None
        if pd.isna(col_a):
            continue
        if pd.isna(col_b):
            # A category header row: only column A has a value.
            if col_a != "Name":
                current_category = col_a
            continue
        if col_a == "Name":
            continue
        records.append({"category": current_category, "master_name": col_a, "price": col_b})

    df = pd.DataFrame.from_records(records)
    df["master_name"] = df["master_name"].astype(str).str.strip()
    return df


# --------------------------------------------------------------------------
# 2. LOAD ONE MONTH OF "ยอดขายตามสินค้า" (sales-by-product) POS export
# --------------------------------------------------------------------------

def load_pos_monthly_sales(path: str | Path, sheet_name=0) -> pd.DataFrame:
    """Read a monthly sales-by-product export into [pos_name, qty, net_sales].

    Robust to extra columns (some months' exports add ช่องทาง / กำไร columns)
    by looking columns up by their Thai header name rather than position.
    """
    df = pd.read_excel(path, sheet_name=sheet_name)

    name_col = "ชื่อสินค้า"
    qty_col = "จำนวน"
    net_col = "ยอดขายสุทธิ (฿)"
    for col in (name_col, qty_col, net_col):
        if col not in df.columns:
            raise ValueError(f"{path}: expected column '{col}' not found. Columns seen: {list(df.columns)}")

    out = df[[name_col, net_col, qty_col]].copy()
    out.columns = ["pos_name", "net_sales", "qty"]

    # Drop the grand-total row and any fully-blank rows.
    out = out[out["pos_name"].notna()]
    out = out[out["pos_name"] != "รวมทั้งหมด"]

    out["net_sales"] = pd.to_numeric(out["net_sales"], errors="coerce").fillna(0.0)
    out["qty"] = pd.to_numeric(out["qty"], errors="coerce").fillna(0.0)
    out["pos_name"] = out["pos_name"].astype(str).str.strip()

    # A given product should appear once per export, but if the POS ever
    # exports duplicate rows for the same name, sum them defensively.
    out = out.groupby("pos_name", as_index=False).sum(numeric_only=True)
    return out


# --------------------------------------------------------------------------
# 3. LOAD RECEIPT-LEVEL DETAIL for cross-checking (optional but recommended)
# --------------------------------------------------------------------------

def load_receipt_detail_totals(path: str | Path, sheet_name: str = "รายละเอียดสินค้าในบิล") -> pd.DataFrame:
    """Aggregate the line-item bill detail sheet into [pos_name, qty, net_sales]
    so it can be compared against the sales-by-product export for the same
    month. This sheet's exact column set has been observed to differ between
    monthly exports (e.g. some months have a 'ยอดขายสุทธิ' column, others
    only 'ยอดขาย (฿)' + separate 'ส่วนลด (฿)'), so columns are looked up by
    header text, in priority order, rather than by fixed position.
    """
    df = pd.read_excel(path, sheet_name=sheet_name)
    cols = list(df.columns)

    def find_col(candidates_in_priority_order, exclude_substrings=()):
        for wanted in candidates_in_priority_order:
            for c in cols:
                s = str(c)
                if wanted in s and not any(bad in s for bad in exclude_substrings):
                    return c
        return None

    name_col = find_col(["ชื่อสินค้า", "รายการ"])
    qty_col = find_col(["จำนวน"])
    # Prefer an explicit net-sales column; otherwise fall back to a generic
    # "ยอดขาย" column (but never อัตรา/ส่วนลด/ภาษี, which are rates/discounts/tax).
    net_col = find_col(
        ["ยอดขายสุทธิ", "สุทธิ", "ยอดขาย"],
        exclude_substrings=["อัตรา", "ส่วนลด", "ภาษี"],
    )

    missing = [n for n, c in [("ชื่อสินค้า", name_col), ("จำนวน", qty_col), ("ยอดขาย", net_col)] if c is None]
    if missing:
        raise ValueError(f"{path}: could not find column(s) {missing} in headers {cols}")

    out = df[[name_col, qty_col, net_col]].copy()
    out.columns = ["pos_name", "qty", "net_sales"]
    out = out[out["pos_name"].notna()]
    # Some exports append a daily-totals block below the real line items,
    # where every text column is literally "\". Drop that junk block.
    out = out[out["pos_name"] != "\\"]
    out["qty"] = pd.to_numeric(out["qty"], errors="coerce").fillna(0.0)
    out["net_sales"] = pd.to_numeric(out["net_sales"], errors="coerce").fillna(0.0)
    out["pos_name"] = out["pos_name"].astype(str).str.strip()

    return out.groupby("pos_name", as_index=False).sum(numeric_only=True)


def cross_check_month(sales_df: pd.DataFrame, receipt_df: pd.DataFrame, month_label: str,
                       qty_tolerance: float = 5.0) -> pd.DataFrame:
    """Compare sales-by-product totals vs. receipt-detail totals for one month.
    Returns rows where quantities disagree by more than qty_tolerance units,
    for a human to review (discounts / voided items / refunds typically
    explain small gaps; large gaps mean a mapping or export problem)."""
    merged = sales_df.merge(receipt_df, on="pos_name", how="outer",
                             suffixes=("_sales_report", "_receipt_detail")).fillna(0.0)
    merged["qty_diff"] = (merged["qty_sales_report"] - merged["qty_receipt_detail"]).abs()
    mismatches = merged[merged["qty_diff"] > qty_tolerance].copy()
    mismatches.insert(0, "month", month_label)
    return mismatches.sort_values("qty_diff", ascending=False)


# --------------------------------------------------------------------------
# 4. APPLY THE NAME MAPPING (POS name -> category + master menu name)
# --------------------------------------------------------------------------

def load_mapping(path: str | Path) -> dict:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def apply_mapping(sales_df: pd.DataFrame, mapping: dict, category_master: pd.DataFrame):
    """Return (mapped_df, unmapped_names).

    mapped_df has columns [pos_name, category, master_name, qty, net_sales].
    unmapped_names is a list of POS names with no mapping entry and no exact
    match to a master menu name either -> needs a human to add to the
    mapping file.
    """
    known_master_names = set(category_master["master_name"])

    categories, master_names, unmapped = [], [], []
    for pos_name in sales_df["pos_name"]:
        if pos_name in mapping:
            categories.append(mapping[pos_name]["category"])
            master_names.append(mapping[pos_name]["master_name"])
        elif pos_name in known_master_names:
            # The POS name already matches a master menu name exactly.
            cat = category_master.loc[category_master["master_name"] == pos_name, "category"].iloc[0]
            categories.append(cat)
            master_names.append(pos_name)
        else:
            categories.append(None)
            master_names.append(None)
            unmapped.append(pos_name)

    out = sales_df.copy()
    out["category"] = categories
    out["master_name"] = master_names
    return out, sorted(set(unmapped))


# --------------------------------------------------------------------------
# 5. BUILD THE FINAL MULTI-MONTH MENU REPORT
# --------------------------------------------------------------------------

def build_menu_report(category_master: pd.DataFrame,
                       monthly_mapped: dict[str, pd.DataFrame],
                       category_filter: str | None = None) -> pd.DataFrame:
    """
    monthly_mapped: {month_label: mapped_df} where mapped_df has columns
                     [master_name, category, qty, net_sales] (already grouped
                     by master_name -- see main()).
    Returns a wide table: one row per master menu item, one qty/net column
    pair per month, totals, rank and status. Items with zero sales in every
    month supplied are kept (with 0s) rather than dropped, since the master
    list is the full current menu.
    """
    base = category_master.copy()
    if category_filter:
        base = base[base["category"] == category_filter].copy()

    report = base[["category", "master_name", "price"]].drop_duplicates("master_name").copy()

    month_labels = list(monthly_mapped.keys())
    for month_label in month_labels:
        m = monthly_mapped[month_label].groupby("master_name", as_index=False)[["qty", "net_sales"]].sum()
        m = m.rename(columns={"qty": f"qty_{month_label}", "net_sales": f"net_{month_label}"})
        report = report.merge(m, on="master_name", how="left")

    qty_cols = [f"qty_{m}" for m in month_labels]
    net_cols = [f"net_{m}" for m in month_labels]
    report[qty_cols + net_cols] = report[qty_cols + net_cols].fillna(0.0)

    report["qty_total"] = report[qty_cols].sum(axis=1)
    report["net_total"] = report[net_cols].sum(axis=1)

    report = report.sort_values("qty_total", ascending=False).reset_index(drop=True)
    report.insert(0, "rank", report.index + 1)

    sold_mask = report["qty_total"] > 0
    n_sold = sold_mask.sum()

    def status(row):
        if row["qty_total"] == 0:
            return "ไม่มีการขาย"
        pct_rank = row["rank"] / n_sold
        if pct_rank <= 0.2:
            return "ขายดี"
        if pct_rank > 0.8:
            return "ขายไม่ดี"
        return "ปานกลาง"

    report["status"] = report.apply(status, axis=1)

    ordered_cols = ["rank", "category", "master_name", "price"]
    for m in month_labels:
        ordered_cols += [f"qty_{m}", f"net_{m}"]
    ordered_cols += ["qty_total", "net_total", "status"]
    return report[ordered_cols]


# --------------------------------------------------------------------------
# 6. CLI GLUE
# --------------------------------------------------------------------------

def _parse_month_args(items: list[str]) -> dict[str, str]:
    out = {}
    for item in items:
        label, _, path = item.partition("=")
        if not path:
            raise ValueError(f"Expected '<month_label>=<path>', got: {item!r}")
        out[label] = path
    return out


def run_pipeline(category_file: str, mapping_file: str, sales_args: list[str],
                  receipts_args: list[str] | None, category_filter: str | None,
                  output_path: str):
    category_master = load_category_master(category_file)
    mapping = load_mapping(mapping_file)

    sales_paths = _parse_month_args(sales_args)
    receipt_paths = _parse_month_args(receipts_args) if receipts_args else {}

    monthly_mapped = {}
    all_unmapped = {}
    for month_label, path in sales_paths.items():
        raw_sales = load_pos_monthly_sales(path)
        mapped, unmapped = apply_mapping(raw_sales, mapping, category_master)
        if unmapped:
            all_unmapped[month_label] = unmapped
        monthly_mapped[month_label] = mapped.dropna(subset=["master_name"])

        if month_label in receipt_paths:
            receipt_totals = load_receipt_detail_totals(receipt_paths[month_label])
            mism = cross_check_month(raw_sales, receipt_totals, month_label)
            if not mism.empty:
                print(f"\n[cross-check] เดือน {month_label}: พบ {len(mism)} รายการที่ยอดจาก "
                      f"'ยอดขายตามสินค้า' กับ 'ใบเสร็จ' ต่างกันเกิน tolerance:")
                print(mism.to_string(index=False))

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
        unmapped_csv = str(Path(output_path).with_suffix("")) + ".unmapped.csv"
        unmapped_flat.to_csv(unmapped_csv, index=False, encoding="utf-8-sig")
        print(f"\n(รายการนี้ถูกบันทึกไว้ที่ {unmapped_csv} ด้วย)")

    report = build_menu_report(category_master, monthly_mapped, category_filter=category_filter)
    report.to_excel(output_path, index=False)
    print(f"\nบันทึกรายงานแล้ว: {output_path}  ({len(report)} เมนู)")
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--category-file", required=True, help="Path to Brown_Ale_Cogs.xlsx (or similar master file)")
    parser.add_argument("--mapping-file", required=True, help="Path to name_mapping.json")
    parser.add_argument("--sales", nargs="+", required=True,
                         help="One or more '<month_label>=<path>' sales-by-product exports")
    parser.add_argument("--receipts", nargs="+", default=None,
                         help="Optional matching '<month_label>=<path>' receipt-detail exports, for cross-checking")
    parser.add_argument("--category-filter", default=None, help="Only keep this category, e.g. Food")
    parser.add_argument("--output", required=True, help="Output .xlsx path")
    args = parser.parse_args()

    run_pipeline(
        category_file=args.category_file,
        mapping_file=args.mapping_file,
        sales_args=args.sales,
        receipts_args=args.receipts,
        category_filter=args.category_filter,
        output_path=args.output,
    )


if __name__ == "__main__":
    main()
