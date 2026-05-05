"""
Generates RAM_Estimate_Builder_Template.xlsx
Upload this file to Google Sheets for online use, or open in Excel 365.
"""
import csv
import os
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ── Load descriptions from CSV (data/descriptions.csv) ───────────────────────
CSV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "descriptions.csv")
descriptions = []
with open(CSV_PATH, encoding="utf-8") as f:
    reader = csv.DictReader(f)
    for row in reader:
        desc = row.get("description", "").strip()
        unit = row.get("unit", "").strip() or ""
        if desc:
            descriptions.append((desc, unit))
print(f"Loaded {len(descriptions)} descriptions from CSV")

wb = openpyxl.Workbook()

# ═══════════════════════════════════════════════════════════════════
# SHEET 1 — DESCRIPTIONS DB (hidden data source)
# ═══════════════════════════════════════════════════════════════════
ws_db = wb.active
ws_db.title = "_DB"
ws_db['A1'] = 'Description'
ws_db['B1'] = 'Unit'
for i, (desc, unit) in enumerate(descriptions, 2):
    ws_db.cell(row=i, column=1, value=desc)
    ws_db.cell(row=i, column=2, value=unit or "")
# keep _DB visible so FILTER formula can find it in both Excel & Google Sheets
DB_ROWS = len(descriptions) + 1    # last row with data

# ═══════════════════════════════════════════════════════════════════
# SHEET 2 — SEARCH (type keywords → see results)
# ═══════════════════════════════════════════════════════════════════
ws_s = wb.create_sheet("🔍 Search")

thin   = Side(style="thin")
ba     = Border(left=thin, right=thin, top=thin, bottom=thin)
hfont  = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
bfont  = Font(name="Calibri", size=10)
hfill  = PatternFill("solid", fgColor="1A3A6E")
sfill  = PatternFill("solid", fgColor="E8F0FB")
yell   = PatternFill("solid", fgColor="FFF2CC")
green  = PatternFill("solid", fgColor="E2EFDA")

def s(ws, r, c, val, font=None, align=None, fill=None, border=None):
    cell = ws.cell(row=r, column=c, value=val)
    if font:   cell.font      = font
    if align:  cell.alignment = align
    if fill:   cell.fill      = fill
    if border: cell.border    = border
    return cell

ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)
lft = Alignment(horizontal="left",   vertical="center", wrap_text=True)
rgt = Alignment(horizontal="right",  vertical="center")

# Title
ws_s.merge_cells("A1:C1")
s(ws_s, 1, 1, "RAM ESTIMATE BUILDER — DESCRIPTION SEARCH",
  Font(name="Calibri", bold=True, size=14, color="1A3A6E"), ctr)
ws_s.row_dimensions[1].height = 30

# Instruction
ws_s.merge_cells("A2:C2")
s(ws_s, 2, 1,
  "HOW TO SEARCH:  Click the dropdown arrow on 'Matching Descriptions' header  ->  type keyword in search box  ->  click OK",
  Font(name="Calibri", size=9, italic=True, color="555555"), lft)
ws_s.row_dimensions[2].height = 18

# Step-by-step tip
ws_s.merge_cells("A3:C3")
s(ws_s, 3, 1,
  "Example keywords:  plastering   |   cement   |   brick   |   excavation   |   RCC   |   painting   |   flooring",
  Font(name="Calibri", size=9, italic=True, color="777777"), lft)
ws_s.row_dimensions[3].height = 16

# Extra tip
ws_s.merge_cells("A4:C4")
s(ws_s, 4, 1,
  "After filtering: select the description -> Ctrl+C to copy -> go to Estimate sheet -> paste in the yellow cell",
  Font(name="Calibri", size=9, italic=True, color="198754"), lft)
ws_s.row_dimensions[4].height = 16

ws_s.row_dimensions[5].height = 6

# Results header — AutoFilter will be enabled on this row
s(ws_s, 6, 1, "#",                    hfont, ctr, hfill, ba)
s(ws_s, 6, 2, "Matching Descriptions",hfont, lft, hfill, ba)
s(ws_s, 6, 3, "Unit",                 hfont, ctr, hfill, ba)
ws_s.row_dimensions[6].height = 22

# ── Write ALL descriptions as STATIC VALUES — no formulas, works everywhere ──
# User uses AutoFilter (the dropdown arrow on row 6) to search by keyword.
# Click the arrow on "Matching Descriptions" → type keyword in search box → all matches show.
for i, (desc, unit) in enumerate(descriptions, 1):
    r = 6 + i
    fill_color = "FFFFFF" if i % 2 == 0 else "F0F6FF"

    # Col A — row number
    an = ws_s.cell(row=r, column=1, value=i)
    an.font      = Font(name="Calibri", size=8, color="999999")
    an.alignment = ctr
    an.border    = ba
    an.fill      = PatternFill("solid", fgColor="F0F0F0")

    # Col B — description (static value)
    bc = ws_s.cell(row=r, column=2, value=desc)
    bc.font      = Font(name="Calibri", size=9)
    bc.alignment = lft
    bc.border    = ba
    bc.fill      = PatternFill("solid", fgColor=fill_color)

    # Col C — unit (static value)
    cc = ws_s.cell(row=r, column=3, value=unit)
    cc.font      = Font(name="Calibri", size=9, color="1A5E9A")
    cc.alignment = ctr
    cc.border    = ba
    cc.fill      = PatternFill("solid", fgColor="EBF5FB" if i % 2 == 0 else "DFF0FF")

    ws_s.row_dimensions[r].height = 26

# Enable AutoFilter on header row so user can search/filter descriptions
ws_s.auto_filter.ref = f"A6:C{6 + len(descriptions)}"

# Freeze rows 1-6 so header stays visible while scrolling
ws_s.freeze_panes = "A7"

# Column widths
ws_s.column_dimensions['A'].width = 5
ws_s.column_dimensions['B'].width = 95
ws_s.column_dimensions['C'].width = 14

# ═══════════════════════════════════════════════════════════════════
# SHEET 3 — ESTIMATE BUILDER
# ═══════════════════════════════════════════════════════════════════
ws_e = wb.create_sheet("📋 Estimate")

# Column widths matching original format
for col, w in [(1,6),(2,40),(3,10),(4,14),(5,10),(6,10),(7,12),(8,12),(9,14)]:
    ws_e.column_dimensions[get_column_letter(col)].width = w

# ── Header ────────────────────────────────────────────────────────
ws_e.merge_cells("A1:I1")
s(ws_e,1,1,"RAM",Font(name="Arial",bold=True,size=11),ctr)
ws_e.row_dimensions[1].height = 14

ws_e.merge_cells("A2:I2")
s(ws_e,2,1,"Detailed Estimate",Font(name="Arial",bold=True,size=12),ctr)
ws_e.row_dimensions[2].height = 18

ws_e.merge_cells("A3:I3")
s(ws_e,3,1,'NAME OF WORK :-  ',Font(name="Arial",bold=True,size=10),lft)
ws_e.row_dimensions[3].height = 24

# ── Column headers ────────────────────────────────────────────────
grey = PatternFill("solid", fgColor="D9D9D9")
ws_e.merge_cells("A4:A5"); ws_e.merge_cells("B4:B5"); ws_e.merge_cells("C4:C5")
ws_e.merge_cells("D4:F4"); ws_e.merge_cells("G4:G5"); ws_e.merge_cells("H4:H5"); ws_e.merge_cells("I4:I5")
for col, h in [(1,"SL. NO."),(2,"Particulars"),(3,"No"),(4,"Measurement"),(7,"Quantity"),(8,"Rate"),(9,"Amount")]:
    s(ws_e,4,col,h,Font(name="Arial",bold=True,size=9),ctr,grey,ba)
ws_e.row_dimensions[4].height = 18
for col,h in [(4,"L"),(5,"B"),(6,"H")]:
    s(ws_e,5,col,h,Font(name="Arial",bold=True,size=9),ctr,grey,ba)
ws_e.row_dimensions[5].height = 14

# Column numbers row
for col in range(1,10):
    s(ws_e,6,col,col,Font(name="Arial",bold=True,size=9),ctr,grey,ba)
ws_e.row_dimensions[6].height = 13

# ── 10 Work Item blocks ───────────────────────────────────────────
START_ROW = 7
MEAS_ROWS = 5        # measurement rows per item
ITEMS     = 10
amount_cells = []

row = START_ROW
for sl in range(1, ITEMS + 1):

    # Description row — YELLOW input cells
    desc_row = row
    s(ws_e, row, 1, sl, Font(name="Arial",bold=True,size=10), ctr, yell, ba)
    ws_e.merge_cells(f"B{row}:E{row}")
    dc = ws_e.cell(row=row, column=2, value=f"← Paste description from Search sheet (Item {sl})")
    dc.font      = Font(name="Arial", size=9, italic=True, color="999999")
    dc.alignment = lft
    dc.border    = ba
    dc.fill      = yell
    for c in range(6, 10):
        s(ws_e, row, c, None, bfont, ctr, None, ba)
    ws_e.row_dimensions[row].height = 42
    row += 1

    # Measurement rows
    qty_cells = []
    for m in range(MEAS_ROWS):
        mr = row
        s(ws_e,mr,1,None,bfont,ctr,None,ba)
        s(ws_e,mr,2,None,bfont,lft,sfill,ba)        # label
        s(ws_e,mr,3,"1",bfont,ctr,None,ba)          # No  (default 1)
        s(ws_e,mr,4,None,bfont,ctr,None,ba)          # L
        s(ws_e,mr,5,None,bfont,ctr,None,ba)          # B
        s(ws_e,mr,6,None,bfont,ctr,None,ba)          # H
        # Qty = No × L × B × H  (skip blanks)
        ws_e.cell(row=mr, column=7,
            value=f"=IF(AND(D{mr}=\"\",E{mr}=\"\",F{mr}=\"\"),\"\","
                  f"IFERROR(C{mr}*IF(D{mr}=\"\",1,D{mr})*IF(E{mr}=\"\",1,E{mr})*IF(F{mr}=\"\",1,F{mr}),\"\"))"
        ).font = Font(name="Arial", size=9)
        ws_e.cell(row=mr, column=7).alignment = ctr
        ws_e.cell(row=mr, column=7).border    = ba
        ws_e.cell(row=mr, column=7).fill      = PatternFill("solid", fgColor="EBF5FB")
        ws_e.cell(row=mr, column=7).number_format = "#,##0.000"
        s(ws_e,mr,8,None,bfont,ctr,None,ba)
        s(ws_e,mr,9,None,bfont,ctr,None,ba)
        ws_e.row_dimensions[mr].height = 16
        qty_cells.append(f"G{mr}")
        row += 1

    # Total / Rate / Amount row
    tr = row
    total_fill = PatternFill("solid", fgColor="FFF2CC")
    bfont_b = Font(name="Arial", bold=True, size=9)
    for c in range(1, 7):
        s(ws_e,tr,c,None,bfont_b,ctr,total_fill,ba)
    ws_e.cell(row=tr, column=7, value=f"=SUM({','.join(qty_cells)})").number_format = "#,##0.000"
    ws_e.cell(row=tr, column=7).font = bfont_b; ws_e.cell(row=tr, column=7).alignment = ctr
    ws_e.cell(row=tr, column=7).fill = total_fill; ws_e.cell(row=tr, column=7).border = ba
    # Rate input cell (H)
    rc = ws_e.cell(row=tr, column=8, value=0)
    rc.font = bfont_b; rc.alignment = rgt; rc.fill = yell; rc.border = ba; rc.number_format = "#,##0.00"
    # Divisor (hidden helper in col — store in comment or use a small helper col)
    # Unit dropdown input
    uc = ws_e.cell(row=tr, column=6, value="1m³")    # unit label
    uc.font = Font(name="Arial",size=8,italic=True); uc.alignment = ctr
    uc.fill = total_fill; uc.border = ba
    # Amount = Qty × Rate
    ac = ws_e.cell(row=tr, column=9,
        value=f"=IFERROR(G{tr}*H{tr},0)")
    ac.font = bfont_b; ac.alignment = rgt
    ac.fill = green; ac.border = ba; ac.number_format = "#,##0.00"
    ws_e.row_dimensions[tr].height = 18
    amount_cells.append(f"I{tr}")
    row += 1

    # Unit label row
    ur = row
    for c in range(1,10): s(ws_e,ur,c,None,bfont,ctr,None,ba)
    ws_e.row_dimensions[ur].height = 11
    row += 1

# ── Footer ────────────────────────────────────────────────────────
footer_row_idx = [row]

def footer(label, formula, fill_color="FFFFFF"):
    r = footer_row_idx[0]
    ws_e.merge_cells(f"A{r}:H{r}")
    ws_e.cell(row=r,column=1,value=label).font      = Font(name="Arial",bold=True,size=10)
    ws_e.cell(row=r,column=1).alignment = rgt
    ws_e.cell(row=r,column=1).fill    = PatternFill("solid",fgColor=fill_color)
    ws_e.cell(row=r,column=1).border  = ba
    ws_e.cell(row=r,column=9,value=formula).font     = Font(name="Arial",bold=True,size=10)
    ws_e.cell(row=r,column=9).alignment = rgt
    ws_e.cell(row=r,column=9).fill    = PatternFill("solid",fgColor=fill_color)
    ws_e.cell(row=r,column=9).border  = ba
    ws_e.cell(row=r,column=9).number_format = "#,##0.00"
    ws_e.row_dimensions[r].height = 18
    footer_row_idx[0] += 1
    return f"I{r}"

sub_ref = footer("Sub Total",        f"=SUM({','.join(amount_cells)})", "D9D9D9")
sei_ref = footer("Add: Seigniorage", 0)
tot_ref = footer("Total",            f"={sub_ref}+{sei_ref}", "FCE4D6")
nac_ref = footer("Add NAC 0.1%",     f"=ROUND({tot_ref}*0.001,0)")
t2_ref  = footer("Total",            f"={tot_ref}+{nac_ref}", "FCE4D6")
gst_ref = footer("Add GST 18%",      f"=ROUND({t2_ref}*0.18,0)")
gt_ref  = footer("Grand Total",      f"={t2_ref}+{gst_ref}", "FCE4D6")

r = footer_row_idx[0]
ws_e.merge_cells(f"A{r}:H{r}")
ws_e.cell(row=r,column=1,value="Say").font      = Font(name="Arial",bold=True,size=11)
ws_e.cell(row=r,column=1).alignment = rgt
ws_e.cell(row=r,column=1).fill   = PatternFill("solid",fgColor="E2EFDA"); ws_e.cell(row=r,column=1).border = ba
ws_e.cell(row=r,column=9,value=f"=ROUND({gt_ref},-2)").font = Font(name="Arial",bold=True,size=11)
ws_e.cell(row=r,column=9).alignment = rgt
ws_e.cell(row=r,column=9).fill   = PatternFill("solid",fgColor="E2EFDA"); ws_e.cell(row=r,column=9).border = ba
ws_e.cell(row=r,column=9).number_format = "#,##0.00"
ws_e.row_dimensions[r].height = 22

# ── Page setup ────────────────────────────────────────────────────
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.page import PageMargins
ws_e.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
ws_e.page_setup.paperSize  = 9
ws_e.page_setup.orientation = "portrait"
ws_e.page_setup.fitToWidth  = 1
ws_e.page_setup.fitToHeight = 0
ws_e.page_margins = PageMargins(left=0.5,right=0.5,top=0.75,bottom=0.75)
ws_e.print_title_rows = "1:6"

# ═══════════════════════════════════════════════════════════════════
# SHEET 4 — INSTRUCTIONS
# ═══════════════════════════════════════════════════════════════════
ws_i = wb.create_sheet("ℹ️ How To Use")
ws_i.column_dimensions['A'].width = 5
ws_i.column_dimensions['B'].width = 90

instructions = [
    ("", "RAM ESTIMATE BUILDER — HOW TO USE", True, "1A3A6E", 14),
    ("", "", False, None, 10),
    ("1️⃣", "SEARCH FOR A DESCRIPTION", True, "2D6FBD", 11),
    ("",  "  • Click the '🔍 Search' tab at the bottom", False, None, 10),
    ("",  "  • Type one or more keywords in cell B4  (e.g.: CC pavement, plastering, excavation)", False, None, 10),
    ("",  "  • Matching descriptions appear in the list below", False, None, 10),
    ("",  "  • Click on a matching description → COPY it (Ctrl+C)", False, None, 10),
    ("", "", False, None, 10),
    ("2️⃣", "ADD THE DESCRIPTION TO THE ESTIMATE", True, "2D6FBD", 11),
    ("",  "  • Click the '📋 Estimate' tab", False, None, 10),
    ("",  "  • Click on the yellow 'Paste description' cell for the item you want", False, None, 10),
    ("",  "  • Paste (Ctrl+V) — description fills in", False, None, 10),
    ("", "", False, None, 10),
    ("3️⃣", "ENTER MEASUREMENTS", True, "2D6FBD", 11),
    ("",  "  • In the rows below the description, enter Label, No, L, B, H", False, None, 10),
    ("",  "  • Qty = No × L × B × H  (auto-calculated — blue cells)", False, None, 10),
    ("",  "  • Use up to 5 measurement rows per item", False, None, 10),
    ("", "", False, None, 10),
    ("4️⃣", "ENTER RATE", True, "2D6FBD", 11),
    ("",  "  • In the yellow RATE cell (column H of the total row), enter the rate from SoR", False, None, 10),
    ("",  "  • Amount = Total Qty × Rate  (auto-calculated — green cell)", False, None, 10),
    ("", "", False, None, 10),
    ("5️⃣", "PRINT / SAVE", True, "2D6FBD", 11),
    ("",  "  • File → Print → already set up for A4 portrait", False, None, 10),
    ("",  "  • Or File → Download → PDF", False, None, 10),
    ("", "", False, None, 10),
    ("💡", "TIPS", True, "198754", 11),
    ("",  "  • Upload this file to Google Sheets for online access from any device", False, None, 10),
    ("",  "  • Share the Google Sheets link with your team — they can all use it", False, None, 10),
    ("",  "  • The _DB sheet is hidden — it contains all 1,590 descriptions", False, None, 10),
]
r = 1
for icon, text, bold, color, size in instructions:
    ws_i.cell(row=r, column=1, value=icon)
    c = ws_i.cell(row=r, column=2, value=text)
    c.font = Font(name="Calibri", bold=bold, size=size,
                  color=color if color else "222222")
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws_i.row_dimensions[r].height = 20 if bold else 18
    r += 1

# ── Set active sheet ──────────────────────────────────────────────
wb.active = ws_e

out = r"C:\Users\STORE-PCS&S\Desktop\rtc\estimate_builder\RAM_Estimate_Template_v3.xlsx"
wb.save(out)
print(f"Template saved: {out}")
print(f"Sheets: {wb.sheetnames}")
print(f"Descriptions loaded: {len(descriptions)}")
print("Done! Open the file and type a keyword in the Search sheet cell B4.")
