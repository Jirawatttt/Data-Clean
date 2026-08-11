"""
Sales Dashboard Builder — ใช้ Sales_JUNE_2569.xlsx เป็น template
=================================================================
วิธีใช้:
  1. แก้ CONFIG ด้านล่าง
  2. วางไฟล์ข้อมูลทั้งหมดในโฟลเดอร์เดียวกัน
  3. รัน: python3 clean_and_build_dashboard_v3.py
"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border
import re, shutil, sys, os

# ============================================================
# CONFIG — แก้ตรงนี้เพื่อใช้กับเดือนใหม่
# ============================================================
MONTH_LABEL   = 'มิถุนายน 2569'        # ชื่อเดือนภาษาไทย
MONTH_SHORT   = 'มิ.ย.'                 # ย่อเดือน
DATE_RANGE    = '01-30 มิ.ย.'           # ช่วงวันที่
TOTAL_DAYS    = 30                       # จำนวนวันในเดือน (เดือนที่ลงท้าย ยน = 30)
OUTPUT_FILE   = 'sales_dashboard_june_2569.xlsx'

FILE_DAILY    = 'D30.xlsx' #รายวัน
FILE_CATEGORY = 'C30.xlsx' #หมวดหมู่
FILE_PRODUCTS = 'P30.xlsx' #สินค้า
FILE_HOURLY   = 'H30.xlsx' #ชั่วโมง
FILE_RECEIPTS = 'R30.xlsx' #ตามบิล
TEMPLATE_FILE = './Template/Template30day.xlsx'

# ============================================================
# Helpers
# ============================================================
def parse_thb(val):
    if pd.isna(val): return 0.0
    return float(str(val).replace('฿','').replace(',','').strip() or 0)

def fmt_hour(s):
    m = re.match(r'(\d+):00-(\d+):00', str(s))
    return f'{int(m.group(1)):02d}-{int(m.group(2)):02d}' if m else str(s)

C_DARK='FF1B3A5C'; C_MED='FF2E6DA4'; C_LIGHT='FF4BACC6'
C_WHITE='FFFFFFFF'; C_STRIPE='FFDDEEFF'; C_BLACK='FF1A1A2E'
NO_BORDER = Border()

def mf(c=C_BLACK, b=False, s=11): return Font(name='Arial', color=c, bold=b, size=s)
def mfill(h): return PatternFill('solid', fgColor=h)
def mal(h='center', v='center', w=False): return Alignment(horizontal=h, vertical=v, wrap_text=w)

# ============================================================
# STEP 1: ตรวจสอบไฟล์
# ============================================================
print("\n📂 โหลดไฟล์ข้อมูล...")
for path, label in [
    (FILE_DAILY,'ยอดขายรายวัน'), (FILE_CATEGORY,'หมวดหมู่'),
    (FILE_PRODUCTS,'สินค้า'), (FILE_HOURLY,'ช่วงเวลา'),
    (FILE_RECEIPTS,'ใบเสร็จ'), (TEMPLATE_FILE,'Template'),
]:
    if not os.path.exists(path):
        print(f"❌ ไม่พบไฟล์ {label}: {path}"); sys.exit(1)
    print(f"   ✅ {label}")

# ============================================================
# STEP 2: คลีน data
# ============================================================

# --- Daily ---
daily_df = pd.read_excel(FILE_DAILY)
daily_df = daily_df[daily_df['วันที่'].astype(str).str.contains('/')].copy()
daily_df['sales']  = daily_df['ยอดขายรวม (฿)'].apply(parse_thb)
daily_df['net']    = daily_df['ยอดขายสุทธิ (฿)'].apply(parse_thb)
daily_df['orders'] = pd.to_numeric(daily_df['จำนวนออเดอร์'], errors='coerce').fillna(0).astype(int)
daily_df['disc']   = daily_df['ส่วนลด (฿)'].apply(parse_thb)
day_th = {0:'จันทร์',1:'อังคาร',2:'พุธ',3:'พฤหัส',4:'ศุกร์',5:'เสาร์',6:'อาทิตย์'}
daily_df['dt']       = pd.to_datetime(daily_df['วันที่'], dayfirst=True)
daily_df['day_th']   = daily_df['dt'].dt.dayofweek.map(day_th)
daily_df['date_str'] = daily_df['dt'].dt.strftime('%d/%m')
daily_df = daily_df.reset_index(drop=True)

# --- Category ---
cat_df = pd.read_excel(FILE_CATEGORY)
cat_df = cat_df[~cat_df['ชื่อหมวดหมู่'].isin(['รวมทั้งหมด','temp'])].copy()
cat_df['sales'] = cat_df['ยอดขายรวม (฿)'].apply(parse_thb)
cat_df['qty']   = pd.to_numeric(cat_df['จำนวน'], errors='coerce').fillna(0).astype(int)
cat_df['pct']   = (cat_df['sales'] / cat_df['sales'].sum() * 100).round(1)
cat_df = cat_df.sort_values('sales', ascending=False).reset_index(drop=True)

# --- Products ---
prod_df = pd.read_excel(FILE_PRODUCTS)
prod_df = prod_df[prod_df['ชื่อสินค้า'] != 'รวมทั้งหมด'].copy()
prod_df['sales'] = prod_df['ยอดขายสุทธิ (฿)'].apply(parse_thb)
top10 = prod_df.sort_values('sales', ascending=False).head(10).reset_index(drop=True)

# --- Hourly (8 rows → H7:H14) ---
hour_df = pd.read_excel(FILE_HOURLY)
hour_df = hour_df[hour_df['ระยะเวลา'] != 'รวมทั้งหมด'].copy()
hour_df['sales']    = hour_df['ยอดขายรวม (฿)'].apply(parse_thb)
hour_df['orders']   = pd.to_numeric(hour_df['จำนวนออเดอร์'], errors='coerce').fillna(0).astype(int)
hour_df = hour_df[hour_df['sales'] > 0].reset_index(drop=True)
hour_df['time_str'] = hour_df['ระยะเวลา'].apply(fmt_hour)
hour_df = hour_df.head(8)   # 8 rows เพราะ total อยู่ H15

# --- Receipts ---
rec_df    = pd.read_excel(FILE_RECEIPTS)
num_bills = len(rec_df[rec_df['เลขที่ใบเสร็จ'].astype(str) != '-'])

# --- Stats ---
total_sales  = daily_df['sales'].sum()
total_items  = int(rec_df['จำนวน'].sum())
total_disc   = daily_df['disc'].sum()
avg_per_bill = round(total_sales / num_bills) if num_bills > 0 else 0
n_open       = len(daily_df[daily_df['sales'] > 0])

# row ที่ daily data จบ และ total row
DAILY_START  = 7
DAILY_END    = DAILY_START + TOTAL_DAYS - 1   # row 36
TOTAL_ROW    = DAILY_END + 1                  # row 37
GAP_ROW      = TOTAL_ROW + 1                  # row 38
SEC_ROW      = TOTAL_ROW + 2                  # row 39
HDR_ROW      = SEC_ROW + 1                    # row 40
DATA_START   = HDR_ROW + 1                    # row 41
CAT_ROWS     = 9
TOP10_ROWS   = 10
CAT_TOTAL_R  = DATA_START + CAT_ROWS          # row 50
HOURLY_TOTAL = 15                             # H15

print(f"\n📊 สรุปสถิติ")
print(f"   ยอดขายรวม : {total_sales:,.0f} บาท")
print(f"   บิล       : {num_bills}  |  เปิด {n_open}/{TOTAL_DAYS} วัน")
print(f"   รายการ    : {total_items:,}  |  เฉลี่ย/บิล {avg_per_bill:,} บาท")
print(f"   ส่วนลด   : {total_disc:,.0f} บาท ({round(total_disc/total_sales*100,1) if total_sales else 0:.1f}%)")

# ============================================================
# STEP 3: สร้าง Excel
# ============================================================
print("\n🏗️  สร้าง Dashboard...")

shutil.copy(TEMPLATE_FILE, OUTPUT_FILE)
wb = load_workbook(OUTPUT_FILE)

# ลบ sheet ที่ไม่ต้องการ
for s in ['กราฟวิเคราะห์','วิเคราะห์และแผน']:
    if s in wb.sheetnames: del wb[s]

ws = wb.active

# Unmerge + clear ทุกอย่างรวม border และ number format
for mr in list(ws.merged_cells.ranges): ws.unmerge_cells(str(mr))
for row in ws.iter_rows():
    for cell in row:
        cell.value         = None
        cell.fill          = PatternFill('none')
        cell.border        = NO_BORDER
        cell.number_format = 'General'

# ── Title ──────────────────────────────────────────────────
ws['B1'].value = f'🏪  SALES DASHBOARD  |  ยอดขายประจำเดือน {MONTH_LABEL}  ({DATE_RANGE})'
ws['B1'].fill = mfill(C_DARK); ws['B1'].font = mf(C_WHITE, True, 14)
ws['B1'].alignment = mal(); ws.merge_cells('B1:J1')
ws.row_dimensions[1].height = 45

# ── KPI Cards row 2-3 ──────────────────────────────────────
for coord, txt, fill in [
    ('B2', f'💰 ยอดขายรวม\n{total_sales:,.0f} บาท',  C_MED),
    ('D2', f'📦 ออเดอร์\n{num_bills}',                 C_LIGHT),
    ('H2', f' เฉลี่ย/บิล\n~{avg_per_bill:,} บาท',    C_DARK),
    ('J2', f'🏷️ ส่วนลด\n{total_disc:,.0f} บาท',       C_MED),
]:
    ws[coord].value = txt; ws[coord].fill = mfill(fill)
    ws[coord].font = mf(C_WHITE, True, 11); ws[coord].alignment = mal(w=True)
ws.merge_cells('B2:C3')
ws.merge_cells('D2:F3')
ws.merge_cells('H2:I3')
ws.merge_cells('J2:J3')
ws.row_dimensions[2].height = 49.5
ws.row_dimensions[3].height = 49.5

# ── Section headers row 5 ──────────────────────────────────
ws['B5'].value = '📅  ยอดขายรายวัน'
ws['B5'].fill = mfill(C_MED); ws['B5'].font = mf(C_WHITE, True)
ws['B5'].alignment = mal(); ws.merge_cells('B5:F5')
ws.row_dimensions[5].height = 25.5
ws['H5'].value = '⏰  ยอดขายตามช่วงเวลา'
ws['H5'].fill = mfill(C_MED); ws['H5'].font = mf(C_WHITE, True)
ws['H5'].alignment = mal(); ws.merge_cells('H5:J5')

# ── Column headers row 6 (ไม่มี ฿) ────────────────────────
for col, lbl in [('B','วันที่'),('C','วัน'),('D','ยอดขาย'),('E','ออเดอ'),('F','ยอดสุทธิ')]:
    c = ws[f'{col}6']; c.value = lbl
    c.fill = mfill(C_DARK); c.font = mf(C_WHITE, True); c.alignment = mal()
for col, lbl in [('H','ช่วงเวลา'),('I','ยอดขาย'),('J','ออเดอ')]:
    c = ws[f'{col}6']; c.value = lbl
    c.fill = mfill(C_DARK); c.font = mf(C_WHITE, True); c.alignment = mal()
ws.row_dimensions[6].height = 19.5

# Column widths
for col, w in [('A',2),('B',20),('C',14),('D',14),('E',8),('F',14),('G',3),('H',22),('I',14),('J',10)]:
    ws.column_dimensions[col].width = w

# ── Daily data rows 7-36 (30 วัน) ─────────────────────────
for day_idx in range(TOTAL_DAYS):
    r = DAILY_START + day_idx
    fill = C_STRIPE if day_idx % 2 == 0 else C_WHITE
    ws.row_dimensions[r].height = 19.5
    if day_idx < len(daily_df):
        row = daily_df.iloc[day_idx]
        vals = [('B', row['date_str']), ('C', row['day_th']),
                ('D', round(row['sales'])), ('E', int(row['orders'])), ('F', round(row['net']))]
    else:
        vals = [('B',''),('C',''),('D',0),('E',0),('F',0)]
    for col, val in vals:
        c = ws[f'{col}{r}']; c.value = val
        c.fill = mfill(fill); c.font = mf(C_BLACK); c.alignment = mal()
        c.border = NO_BORDER

# ── Hourly data rows 7-14 (8 rows, ไม่มีเส้น) ─────────────
for i, row in hour_df.iterrows():
    r = DAILY_START + i
    fill = C_STRIPE if i % 2 == 0 else C_WHITE
    for col, val in [('H', row['time_str']), ('I', int(row['sales'])), ('J', int(row['orders']))]:
        c = ws[f'{col}{r}']; c.value = val
        c.fill = mfill(fill); c.font = mf(C_BLACK); c.alignment = mal()
        c.border = NO_BORDER

# ── Hourly total row 15 ────────────────────────────────────
ws[f'H{HOURLY_TOTAL}'].value = 'รวม'
ws[f'H{HOURLY_TOTAL}'].fill = mfill(C_MED); ws[f'H{HOURLY_TOTAL}'].font = mf(C_WHITE, True)
ws[f'H{HOURLY_TOTAL}'].alignment = mal()
ws[f'I{HOURLY_TOTAL}'].value = f'=SUM(I{DAILY_START}:I{DAILY_START+7})'
ws[f'I{HOURLY_TOTAL}'].fill = mfill(C_MED); ws[f'I{HOURLY_TOTAL}'].font = mf(C_WHITE, True)
ws[f'I{HOURLY_TOTAL}'].alignment = mal()
ws[f'J{HOURLY_TOTAL}'].value = f'=SUM(J{DAILY_START}:J{DAILY_START+7})'
ws[f'J{HOURLY_TOTAL}'].fill = mfill(C_MED); ws[f'J{HOURLY_TOTAL}'].font = mf(C_WHITE, True)
ws[f'J{HOURLY_TOTAL}'].alignment = mal()
ws.row_dimensions[HOURLY_TOTAL].height = 19.5

# ── Daily total row 37 ─────────────────────────────────────
ws[f'B{TOTAL_ROW}'].value = 'รวม'
ws[f'B{TOTAL_ROW}'].fill = mfill(C_MED); ws[f'B{TOTAL_ROW}'].font = mf(C_WHITE, True)
ws[f'B{TOTAL_ROW}'].alignment = mal()
ws[f'C{TOTAL_ROW}'].value = n_open          # จำนวนวันที่มีการขาย
ws[f'C{TOTAL_ROW}'].fill = mfill(C_MED); ws[f'C{TOTAL_ROW}'].font = mf(C_WHITE, True)
ws[f'C{TOTAL_ROW}'].alignment = mal()
for col in ['D','E','F']:
    c = ws[f'{col}{TOTAL_ROW}']
    c.value = f'=SUM({col}{DAILY_START}:{col}{DAILY_END})'
    c.fill = mfill(C_MED); c.font = mf(C_WHITE, True); c.alignment = mal()
ws.row_dimensions[TOTAL_ROW].height = 19.5

# ── Gap row 38 ─────────────────────────────────────────────
ws.row_dimensions[GAP_ROW].height = 7.5

# ── Section headers row 39 ─────────────────────────────────
ws[f'B{SEC_ROW}'].value = '🗂️  ยอดขายตามหมวดหมู่'
ws[f'B{SEC_ROW}'].fill = mfill(C_MED); ws[f'B{SEC_ROW}'].font = mf(C_WHITE, True)
ws[f'B{SEC_ROW}'].alignment = mal(); ws.merge_cells(f'B{SEC_ROW}:E{SEC_ROW}')
ws[f'H{SEC_ROW}'].value = '🏆  TOP 10 สินค้าขายดี'
ws[f'H{SEC_ROW}'].fill = mfill(C_MED); ws[f'H{SEC_ROW}'].font = mf(C_WHITE, True)
ws[f'H{SEC_ROW}'].alignment = mal(); ws.merge_cells(f'H{SEC_ROW}:J{SEC_ROW}')
ws.row_dimensions[SEC_ROW].height = 19.5

# ── Column headers row 40 (ไม่มี ฿) ───────────────────────
for col, lbl in [('B','หมวดหมู่'),('C','ยอดขาย'),('D','จำนวน'),('E','%')]:
    c = ws[f'{col}{HDR_ROW}']; c.value = lbl
    c.fill = mfill(C_DARK); c.font = mf(C_WHITE, True); c.alignment = mal()
for col, lbl in [('H','สินค้า'),('I','ยอดขาย'),('J','จำนวน')]:
    c = ws[f'{col}{HDR_ROW}']; c.value = lbl
    c.fill = mfill(C_DARK); c.font = mf(C_WHITE, True); c.alignment = mal()
ws.row_dimensions[HDR_ROW].height = 19.5

# ── Category data rows 41-49 (9 rows, E = %) ───────────────
for i in range(CAT_ROWS):
    r = DATA_START + i
    fill = C_STRIPE if i % 2 == 0 else C_WHITE
    ws.row_dimensions[r].height = 19.5
    if i < len(cat_df):
        row = cat_df.iloc[i]
        for col, val in [('B', row['ชื่อหมวดหมู่']), ('C', int(row['sales'])),
                         ('D', row['qty']), ('E', row['pct'])]:
            c = ws[f'{col}{r}']; c.value = val
            c.fill = mfill(fill); c.font = mf(C_BLACK); c.alignment = mal()
    else:
        for col in ['B','C','D','E']:
            c = ws[f'{col}{r}']; c.value = ''; c.fill = mfill(fill)

# ── Top10 data rows 41-50 (10 rows) ────────────────────────
for i in range(TOP10_ROWS):
    r = DATA_START + i
    fill = C_STRIPE if i % 2 == 0 else C_WHITE
    ws.row_dimensions[r].height = 19.5
    if i < len(top10):
        row = top10.iloc[i]
        for col, val in [('H', f'{i+1}. {row["ชื่อสินค้า"]}'),
                         ('I', int(row['sales'])), ('J', int(row['จำนวน']))]:
            c = ws[f'{col}{r}']; c.value = val
            c.fill = mfill(fill); c.font = mf(C_BLACK)
            c.alignment = mal(h='left' if col == 'H' else 'center')

# ── Category total row 50 ──────────────────────────────────
ws[f'B{CAT_TOTAL_R}'].value = 'รวม'
ws[f'B{CAT_TOTAL_R}'].fill = mfill(C_MED); ws[f'B{CAT_TOTAL_R}'].font = mf(C_WHITE, True)
ws[f'B{CAT_TOTAL_R}'].alignment = mal()
for col in ['C','D','E']:
    c = ws[f'{col}{CAT_TOTAL_R}']
    c.value = f'=SUM({col}{DATA_START}:{col}{CAT_TOTAL_R-1})'
    c.fill = mfill(C_MED); c.font = mf(C_WHITE, True); c.alignment = mal()
ws.row_dimensions[CAT_TOTAL_R].height = 19.5

# ============================================================
# STEP 4: Save
# ============================================================
wb.save(OUTPUT_FILE)
print(f"\n✅ บันทึกสำเร็จ → {OUTPUT_FILE}")
print("=" * 50)
