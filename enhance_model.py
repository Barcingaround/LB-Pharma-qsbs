#!/usr/bin/env python3
"""
Comprehensive model enhancement:
1. Fix circular references in Scenarios sheet
2. Add Probability-Weighted EV sheet
3. Add Breakeven Analysis
4. Add After-Tax IRR
5. Add Dilution Waterfall
6. Enhance Sensitivity tables
"""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from copy import copy

wb = load_workbook('/home/user/LB-Pharma-qsbs/LBRX_QSBS_Evecxia_Model.xlsx')

# ============================================================
# STYLE DEFINITIONS
# ============================================================
header_font = Font(name='Calibri', bold=True, size=11)
title_font = Font(name='Calibri', bold=True, size=13)
section_font = Font(name='Calibri', bold=True, size=11, color='FFFFFF')
section_fill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
light_blue_fill = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
light_green_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
light_red_fill = PatternFill(start_color='FCE4EC', end_color='FCE4EC', fill_type='solid')
light_yellow_fill = PatternFill(start_color='FFF9C4', end_color='FFF9C4', fill_type='solid')
input_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
pct_fmt = '0.0%'
dollar_fmt = '$#,##0'
dollar_dec_fmt = '$#,##0.00'
big_dollar_fmt = '$#,##0'
moic_fmt = '0.0x'
number_fmt = '#,##0'

def style_header_row(ws, row, max_col, font=section_font, fill=section_fill):
    for c in range(1, max_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = font
        cell.fill = fill
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = thin_border

def style_range(ws, min_row, max_row, min_col, max_col, fill=None):
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = thin_border
            if fill:
                cell.fill = fill

# ============================================================
# 1. FIX SCENARIOS SHEET — Rebuild the main grid (rows 3-26)
# ============================================================
print("1. Fixing Scenarios sheet formulas...")
ws = wb['Scenarios']

# Clear the problematic rows 3-26 and rebuild with correct formulas
# Row layout:
# 3: Headers (already there, keep)
# 4: Sale Date
# 5: Sale Price / Share (LBRX) or Company Val (Evecxia)
# 6: Shares Sold / Owned
# 7: Gross Proceeds
# 8: Cost Basis
# 9: Gross Gain
# 10: Holding Period (days)
# 11: Holding Period (years)
# 12: Long-Term?
# 13: QSBS 5-Year Met?
# 14: §1202 Exclusion Amount
# 15: §1045 Deferral
# 16: Taxable Gain
# 17: Federal Tax
# 18: State Tax
# 19: Total Tax
# 20: Net After-Tax
# 21: MOIC
# 22: Effective Tax Rate

# Column mapping: B=S1(Lock-Up), C=S2(Ph3+), D=S3(Ph3-), E=S4(Dual), F=S5(Eve+), G=S6(Eve-)

# Row 3: Section header - keep
# Row 4: Scenario labels
labels = [
    (5, 'Sale Date'),
    (6, 'Sale Price / Share'),
    (7, 'Shares'),
    (8, 'Gross Proceeds'),
    (9, 'Cost Basis'),
    (10, 'Gross Gain'),
    (11, 'Holding Period (days)'),
    (12, 'Holding Period (years)'),
    (13, 'Long-Term?'),
    (14, 'QSBS 5-Year Met?'),
    (15, '§1202 Exclusion'),
    (16, '§1045 Deferral'),
    (17, 'Taxable Gain'),
    (18, 'Federal Tax'),
    (19, 'State Tax'),
    (20, 'Total Tax'),
    (21, 'Net After-Tax'),
    (22, 'MOIC'),
    (23, 'Effective Tax Rate'),
]

for row_num, label in labels:
    ws.cell(row=row_num, column=1, value=label).font = header_font
    ws.cell(row=row_num, column=1).border = thin_border

# S1: Lock-Up Sale (col B)
ws['B5'] = '=lockup_date'
ws['B6'] = '=lockup_price'
ws['B7'] = '=serb_shares'
ws['B8'] = '=B6*B7'
ws['B9'] = '=serb_basis'
ws['B10'] = '=B8-B9'
ws['B11'] = '=lockup_date-serb_date'
ws['B12'] = '=B11/365.25'
ws['B13'] = '=IF(B11>365,"YES","NO")'
ws['B14'] = '=IF(AND(B11>=1826,Inputs!B18="YES"),"YES","NO")'
ws['B15'] = '=IF(B14="YES",MIN(B10,MAX(10000000,10*B9)),0)'
ws['B16'] = 0  # no 1045 deferral for straight sale
ws['B17'] = '=MAX(0,B10-B15-B16)'
ws['B18'] = '=IF(B13="YES",B17*eff_fed_rate,B17*Inputs!B11)'
ws['B19'] = '=B17*state_rate'
ws['B20'] = '=B18+B19'
ws['B21'] = '=B8-B20'
ws['B22'] = '=B21/B9'
ws['B23'] = '=IF(B10>0,B20/B10,0)'

# S2: Ph3+ Hold (col C)
ws['C5'] = '=ph3_date'
ws['C6'] = '=Inputs!B53'  # Ph3+ base price
ws['C7'] = '=serb_shares'
ws['C8'] = '=C6*C7'
ws['C9'] = '=serb_basis'
ws['C10'] = '=C8-C9'
ws['C11'] = '=ph3_date-serb_date'
ws['C12'] = '=C11/365.25'
ws['C13'] = '=IF(C11>365,"YES","NO")'
ws['C14'] = '=IF(AND(C11>=1826,Inputs!B18="YES"),"YES","NO")'
ws['C15'] = '=IF(C14="YES",MIN(C10,MAX(10000000,10*C9)),0)'
ws['C16'] = 0
ws['C17'] = '=MAX(0,C10-C15-C16)'
ws['C18'] = '=IF(C13="YES",C17*eff_fed_rate,C17*Inputs!B11)'
ws['C19'] = '=C17*state_rate'
ws['C20'] = '=C18+C19'
ws['C21'] = '=C8-C20'
ws['C22'] = '=C21/C9'
ws['C23'] = '=IF(C10>0,C20/C10,0)'

# S3: Ph3- Hold (col D)
ws['D5'] = '=ph3_date'
ws['D6'] = '=Inputs!B56'  # Ph3- base price
ws['D7'] = '=serb_shares'
ws['D8'] = '=D6*D7'
ws['D9'] = '=serb_basis'
ws['D10'] = '=D8-D9'
ws['D11'] = '=ph3_date-serb_date'
ws['D12'] = '=D11/365.25'
ws['D13'] = '=IF(D11>365,"YES","NO")'
ws['D14'] = '=IF(AND(D11>=1826,Inputs!B18="YES"),"YES","NO")'
ws['D15'] = '=IF(D14="YES",MIN(MAX(0,D10),MAX(10000000,10*D9)),0)'
ws['D16'] = 0
ws['D17'] = '=MAX(0,D10-D15-D16)'
ws['D18'] = '=IF(D13="YES",D17*eff_fed_rate,D17*Inputs!B11)'
ws['D19'] = '=D17*state_rate'
ws['D20'] = '=D18+D19'
ws['D21'] = '=D8-D20'
ws['D22'] = '=D21/D9'
ws['D23'] = '=IF(D10>0,D20/D10,0)'

# S4: Dual Lot Lock-Up (col E) — simplified as combined
ws['E5'] = '=lockup_date'
ws['E6'] = '=lockup_price'
ws['E7'] = '=serb_shares+IF(Inputs!B23="",0,Inputs!B25)'
ws['E8'] = '=E6*E7'
ws['E9'] = '=serb_basis+IF(Inputs!B23="",0,Inputs!B23)'
ws['E10'] = '=E8-E9'
ws['E11'] = '=lockup_date-serb_date'
ws['E12'] = '=E11/365.25'
ws['E13'] = '=IF(E11>365,"YES","NO")'
ws['E14'] = '=IF(AND(E11>=1826,Inputs!B18="YES"),"YES","NO")'
# Only Series B portion gets 1202; Series C does not
ws['E15'] = '=IF(E14="YES",MIN(serb_shares*lockup_price-serb_basis,MAX(10000000,10*serb_basis)),0)'
ws['E16'] = 0
ws['E17'] = '=MAX(0,E10-E15-E16)'
ws['E18'] = '=IF(E13="YES",E17*eff_fed_rate,E17*Inputs!B11)'
ws['E19'] = '=E17*state_rate'
ws['E20'] = '=E18+E19'
ws['E21'] = '=E8-E20'
ws['E22'] = '=E21/E9'
ws['E23'] = '=IF(E10>0,E20/E10,0)'

# S5: §1045 → Evecxia Ph2+ (col F)
ws['F5'] = '=evecxia_ph2_date'
ws['F6'] = '=Inputs!B59/Inputs!B65'  # Ph2+ base valuation / shares = price per share
ws['F7'] = '=Inputs!B32'  # Evecxia shares acquired
ws['F8'] = '=F6*F7'  # Gross proceeds
ws['F9'] = '=serb_basis'  # Original LBRX basis carries through via 1045
ws['F10'] = '=F8-F9'
ws['F11'] = '=evecxia_ph2_date-serb_date'  # Tacked holding from Series B date
ws['F12'] = '=F11/365.25'
ws['F13'] = '=IF(F11>365,"YES","NO")'
ws['F14'] = '=IF(AND(F11>=1826,Inputs!B33="YES"),"YES","NO")'
ws['F15'] = '=IF(F14="YES",MIN(F10,MAX(10000000,10*F9)),0)'
ws['F16'] = '=serb_shares*lockup_price-serb_basis'  # Gain deferred from LBRX sale
ws['F17'] = '=MAX(0,F10-F15)'  # 1045 deferral already baked into basis
ws['F18'] = '=IF(F13="YES",F17*eff_fed_rate,F17*Inputs!B11)'
ws['F19'] = '=F17*state_rate'
ws['F20'] = '=F18+F19'
ws['F21'] = '=F8-F20'
ws['F22'] = '=F21/F9'
ws['F23'] = '=IF(F10>0,F20/F10,0)'

# S6: §1045 → Evecxia Ph2- (col G)
ws['G5'] = '=evecxia_ph2_date'
ws['G6'] = '=Inputs!B62/Inputs!B65'  # Ph2- base valuation / shares
ws['G7'] = '=Inputs!B32'
ws['G8'] = '=G6*G7'
ws['G9'] = '=serb_basis'
ws['G10'] = '=G8-G9'
ws['G11'] = '=evecxia_ph2_date-serb_date'
ws['G12'] = '=G11/365.25'
ws['G13'] = '=IF(G11>365,"YES","NO")'
ws['G14'] = '=IF(AND(G11>=1826,Inputs!B33="YES"),"YES","NO")'
ws['G15'] = '=IF(G14="YES",MIN(MAX(0,G10),MAX(10000000,10*G9)),0)'
ws['G16'] = '=serb_shares*lockup_price-serb_basis'
ws['G17'] = '=MAX(0,G10-G15)'
ws['G18'] = '=IF(G13="YES",G17*eff_fed_rate,G17*Inputs!B11)'
ws['G19'] = '=G17*state_rate'
ws['G20'] = '=G18+G19'
ws['G21'] = '=G8-G20'
ws['G22'] = '=G21/G9'
ws['G23'] = '=IF(G10>0,G20/G10,0)'

# Format Scenarios grid
for row in range(5, 24):
    for col in range(2, 8):
        cell = ws.cell(row=row, column=col)
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center')

# Number formats
for col in range(2, 8):
    ws.cell(row=5, column=col).number_format = 'YYYY-MM-DD'
    ws.cell(row=6, column=col).number_format = dollar_dec_fmt
    ws.cell(row=7, column=col).number_format = number_fmt
    ws.cell(row=8, column=col).number_format = dollar_fmt
    ws.cell(row=9, column=col).number_format = dollar_fmt
    ws.cell(row=10, column=col).number_format = dollar_fmt
    ws.cell(row=11, column=col).number_format = number_fmt
    ws.cell(row=12, column=col).number_format = '0.00'
    ws.cell(row=15, column=col).number_format = dollar_fmt
    ws.cell(row=16, column=col).number_format = dollar_fmt
    ws.cell(row=17, column=col).number_format = dollar_fmt
    ws.cell(row=18, column=col).number_format = dollar_fmt
    ws.cell(row=19, column=col).number_format = dollar_fmt
    ws.cell(row=20, column=col).number_format = dollar_fmt
    ws.cell(row=21, column=col).number_format = dollar_fmt
    ws.cell(row=22, column=col).number_format = '0.0x'
    ws.cell(row=23, column=col).number_format = pct_fmt

# Highlight net row
for col in range(2, 8):
    ws.cell(row=21, column=col).font = Font(bold=True, size=11)
    ws.cell(row=22, column=col).font = Font(bold=True, size=11)

print("   Scenarios sheet fixed.")

# ============================================================
# 2. ADD PROBABILITY-WEIGHTED EV SHEET
# ============================================================
print("2. Adding Decision Matrix (Probability-Weighted EV)...")

if 'Decision Matrix' in wb.sheetnames:
    del wb['Decision Matrix']
ws_dm = wb.create_sheet('Decision Matrix', wb.sheetnames.index('Outputs'))

ws_dm.column_dimensions['A'].width = 36
for c in ['B','C','D','E','F','G','H']:
    ws_dm.column_dimensions[c].width = 18

# Title
ws_dm['A1'] = 'DECISION MATRIX — PROBABILITY-WEIGHTED ANALYSIS'
ws_dm['A1'].font = title_font

# --- Section 1: Clinical Probability Inputs ---
ws_dm['A3'] = 'CLINICAL PROBABILITY ASSUMPTIONS'
style_header_row(ws_dm, 3, 4)

ws_dm['A4'] = 'P(LBRX Phase 3 Success)'
ws_dm['B4'] = 0.55
ws_dm['B4'].number_format = pct_fmt
ws_dm['B4'].fill = input_fill
ws_dm['C4'] = '← Adjustable input (industry avg MDD Ph3: 50-60%)'
ws_dm['C4'].font = Font(italic=True, color='666666', size=10)

ws_dm['A5'] = 'P(LBRX Phase 3 Failure)'
ws_dm['B5'] = '=1-B4'
ws_dm['B5'].number_format = pct_fmt

ws_dm['A7'] = 'P(Evecxia Phase 2 Success)'
ws_dm['B7'] = 0.35
ws_dm['B7'].number_format = pct_fmt
ws_dm['B7'].fill = input_fill
ws_dm['C7'] = '← Adjustable input (CNS Ph2 success rate: 30-40%)'
ws_dm['C7'].font = Font(italic=True, color='666666', size=10)

ws_dm['A8'] = 'P(Evecxia Phase 2 Failure)'
ws_dm['B8'] = '=1-B7'
ws_dm['B8'].number_format = pct_fmt

for r in range(4, 9):
    for c in range(1, 3):
        ws_dm.cell(row=r, column=c).border = thin_border

# --- Section 2: Path 2 — Hold LBRX ---
ws_dm['A10'] = 'PATH 2: HOLD LBRX THROUGH PHASE 3'
style_header_row(ws_dm, 10, 6)

headers_p2 = ['Scenario', 'Probability', 'LBRX Price', 'Gross Proceeds', 'Tax', 'Net After-Tax']
for i, h in enumerate(headers_p2, 1):
    ws_dm.cell(row=11, column=i, value=h).font = header_font
    ws_dm.cell(row=11, column=i).fill = light_blue_fill
    ws_dm.cell(row=11, column=i).border = thin_border
    ws_dm.cell(row=11, column=i).alignment = Alignment(horizontal='center')

# Ph3+ Low
ws_dm['A12'] = 'Ph3 Positive — Low ($50)'
ws_dm['B12'] = '=B4*0.25'  # 25% of positive scenarios
ws_dm['C12'] = '=Inputs!B52'
ws_dm['D12'] = '=C12*serb_shares'
ws_dm['E12'] = 0  # QSBS excluded
ws_dm['F12'] = '=D12-E12'

# Ph3+ Base
ws_dm['A13'] = 'Ph3 Positive — Base ($100)'
ws_dm['B13'] = '=B4*0.50'
ws_dm['C13'] = '=Inputs!B53'
ws_dm['D13'] = '=C13*serb_shares'
ws_dm['E13'] = 0
ws_dm['F13'] = '=D13-E13'

# Ph3+ High
ws_dm['A14'] = 'Ph3 Positive — High ($165)'
ws_dm['B14'] = '=B4*0.25'
ws_dm['C14'] = '=Inputs!B54'
ws_dm['D14'] = '=C14*serb_shares'
ws_dm['E14'] = 0
ws_dm['F14'] = '=D14-E14'

# Ph3- Low
ws_dm['A15'] = 'Ph3 Negative — Low ($3.30)'
ws_dm['B15'] = '=B5*0.25'
ws_dm['C15'] = '=Inputs!B55'
ws_dm['D15'] = '=C15*serb_shares'
ws_dm['E15'] = 0
ws_dm['F15'] = '=D15-E15'

# Ph3- Base
ws_dm['A16'] = 'Ph3 Negative — Base ($6.60)'
ws_dm['B16'] = '=B5*0.50'
ws_dm['C16'] = '=Inputs!B56'
ws_dm['D16'] = '=C16*serb_shares'
ws_dm['E16'] = 0
ws_dm['F16'] = '=D16-E16'

# Ph3- High
ws_dm['A17'] = 'Ph3 Negative — High ($11.50)'
ws_dm['B17'] = '=B5*0.25'
ws_dm['C17'] = '=Inputs!B57'
ws_dm['D17'] = '=C17*serb_shares'
ws_dm['E17'] = 0
ws_dm['F17'] = '=D17-E17'

# Weighted EV
ws_dm['A19'] = 'PATH 2 EXPECTED VALUE'
ws_dm['A19'].font = Font(bold=True, size=12)
ws_dm['B19'] = '=SUM(B12:B17)'
ws_dm['B19'].number_format = pct_fmt
ws_dm['F19'] = '=SUMPRODUCT(B12:B17,F12:F17)/SUM(B12:B17)'
ws_dm['F19'].number_format = dollar_fmt
ws_dm['F19'].font = Font(bold=True, size=12)
ws_dm['A20'] = 'PATH 2 EXPECTED MOIC'
ws_dm['A20'].font = Font(bold=True, size=12)
ws_dm['F20'] = '=F19/serb_basis'
ws_dm['F20'].number_format = '0.00x'
ws_dm['F20'].font = Font(bold=True, size=12)

for r in range(12, 18):
    ws_dm.cell(row=r, column=2).number_format = pct_fmt
    ws_dm.cell(row=r, column=3).number_format = dollar_dec_fmt
    ws_dm.cell(row=r, column=4).number_format = dollar_fmt
    ws_dm.cell(row=r, column=5).number_format = dollar_fmt
    ws_dm.cell(row=r, column=6).number_format = dollar_fmt
    for c in range(1, 7):
        ws_dm.cell(row=r, column=c).border = thin_border

for c in range(1, 7):
    ws_dm.cell(row=19, column=c).border = thin_border
    ws_dm.cell(row=20, column=c).border = thin_border

# --- Section 3: Path 3 — Roll into Evecxia ---
ws_dm['A22'] = 'PATH 3: §1045 ROLL INTO EVECXIA'
style_header_row(ws_dm, 22, 6)

headers_p3 = ['Scenario', 'Probability', 'Evecxia Valuation', 'Gross Proceeds', 'Tax', 'Net After-Tax']
for i, h in enumerate(headers_p3, 1):
    ws_dm.cell(row=23, column=i, value=h).font = header_font
    ws_dm.cell(row=23, column=i).fill = light_green_fill
    ws_dm.cell(row=23, column=i).border = thin_border
    ws_dm.cell(row=23, column=i).alignment = Alignment(horizontal='center')

# Ph2+ Low
ws_dm['A24'] = 'Ph2 Positive — Low ($800M)'
ws_dm['B24'] = '=B7*0.25'
ws_dm['C24'] = '=Inputs!B58'
ws_dm['D24'] = '=C24/Inputs!B65*Inputs!B32'  # valuation / total shares * investor shares
ws_dm['E24'] = 0  # QSBS excluded
ws_dm['F24'] = '=D24-E24'

# Ph2+ Base
ws_dm['A25'] = 'Ph2 Positive — Base ($1.5B)'
ws_dm['B25'] = '=B7*0.50'
ws_dm['C25'] = '=Inputs!B59'
ws_dm['D25'] = '=C25/Inputs!B65*Inputs!B32'
ws_dm['E25'] = 0
ws_dm['F25'] = '=D25-E25'

# Ph2+ High
ws_dm['A26'] = 'Ph2 Positive — High ($3B)'
ws_dm['B26'] = '=B7*0.25'
ws_dm['C26'] = '=Inputs!B60'
ws_dm['D26'] = '=C26/Inputs!B65*Inputs!B32'
ws_dm['E26'] = 0
ws_dm['F26'] = '=D26-E26'

# Ph2- Low
ws_dm['A27'] = 'Ph2 Negative — Low ($50M)'
ws_dm['B27'] = '=B8*0.25'
ws_dm['C27'] = '=Inputs!B61'
ws_dm['D27'] = '=C27/Inputs!B65*Inputs!B32'
ws_dm['E27'] = 0
ws_dm['F27'] = '=D27-E27'

# Ph2- Base
ws_dm['A28'] = 'Ph2 Negative — Base ($120M)'
ws_dm['B28'] = '=B8*0.50'
ws_dm['C28'] = '=Inputs!B62'
ws_dm['D28'] = '=C28/Inputs!B65*Inputs!B32'
ws_dm['E28'] = 0
ws_dm['F28'] = '=D28-E28'

# Ph2- High
ws_dm['A29'] = 'Ph2 Negative — High ($250M)'
ws_dm['B29'] = '=B8*0.25'
ws_dm['C29'] = '=Inputs!B63'
ws_dm['D29'] = '=C29/Inputs!B65*Inputs!B32'
ws_dm['E29'] = 0
ws_dm['F29'] = '=D29-E29'

# Weighted EV
ws_dm['A31'] = 'PATH 3 EXPECTED VALUE'
ws_dm['A31'].font = Font(bold=True, size=12)
ws_dm['B31'] = '=SUM(B24:B29)'
ws_dm['B31'].number_format = pct_fmt
ws_dm['F31'] = '=SUMPRODUCT(B24:B29,F24:F29)/SUM(B24:B29)'
ws_dm['F31'].number_format = dollar_fmt
ws_dm['F31'].font = Font(bold=True, size=12)
ws_dm['A32'] = 'PATH 3 EXPECTED MOIC'
ws_dm['A32'].font = Font(bold=True, size=12)
ws_dm['F32'] = '=F31/serb_basis'
ws_dm['F32'].number_format = '0.00x'
ws_dm['F32'].font = Font(bold=True, size=12)

for r in range(24, 30):
    ws_dm.cell(row=r, column=2).number_format = pct_fmt
    ws_dm.cell(row=r, column=3).number_format = '$#,##0,,\"M\"'
    ws_dm.cell(row=r, column=4).number_format = dollar_fmt
    ws_dm.cell(row=r, column=5).number_format = dollar_fmt
    ws_dm.cell(row=r, column=6).number_format = dollar_fmt
    for c in range(1, 7):
        ws_dm.cell(row=r, column=c).border = thin_border

for c in range(1, 7):
    ws_dm.cell(row=31, column=c).border = thin_border
    ws_dm.cell(row=32, column=c).border = thin_border

# --- Section 4: Head-to-Head Comparison ---
ws_dm['A34'] = 'HEAD-TO-HEAD COMPARISON'
style_header_row(ws_dm, 34, 4)

compare_labels = [
    ('Metric', 'Path 2: Hold LBRX', 'Path 3: Roll to Evecxia', 'Advantage'),
    ('Expected Net Value', '=F19', '=F31', '=IF(C36>B36,"PATH 3","PATH 2")'),
    ('Expected MOIC', '=F20', '=F32', '=IF(C37>B37,"PATH 3","PATH 2")'),
    ('Best Case Net', '=MAX(F12:F17)', '=MAX(F24:F29)', '=IF(C38>B38,"PATH 3","PATH 2")'),
    ('Worst Case Net', '=MIN(F12:F17)', '=MIN(F24:F29)', '=IF(C39>B39,"PATH 2","PATH 3")'),
    ('P(Loss vs Basis)', None, None, None),  # special
    ('Tax at Exit', '$0 (§1202)', '$0 (§1045→§1202)', 'TIED'),
    ('Liquidity', 'Public (post-Ph3)', 'Private→IPO→Lock-up', 'PATH 2'),
    ('EV / Risk Ratio', '=F19/ABS(MIN(F12:F17))', '=F31/ABS(MIN(F24:F29))', '=IF(C43>B43,"PATH 3","PATH 2")'),
]

for i, row_data in enumerate(compare_labels, 35):
    for j, val in enumerate(row_data, 1):
        cell = ws_dm.cell(row=i, column=j)
        if val is not None:
            cell.value = val
        cell.border = thin_border
        if j == 1:
            cell.font = header_font
        if j == 4 and val and 'PATH 3' in str(val):
            cell.font = Font(bold=True, color='006600')

# Header row for comparison
for c in range(1, 5):
    ws_dm.cell(row=35, column=c).font = header_font
    ws_dm.cell(row=35, column=c).fill = light_yellow_fill
    ws_dm.cell(row=35, column=c).alignment = Alignment(horizontal='center')

# Special: P(Loss) — loss = net < basis
ws_dm['A40'] = 'P(Net < Basis)'
ws_dm['B40'] = '=SUMPRODUCT((F12:F17<serb_basis)*B12:B17)'
ws_dm['C40'] = '=SUMPRODUCT((F24:F29<serb_basis)*B24:B29)'
ws_dm['D40'] = '=IF(C40<B40,"PATH 3","PATH 2")'
ws_dm['B40'].number_format = pct_fmt
ws_dm['C40'].number_format = pct_fmt

# Format money cells
for r in [36, 38, 39]:
    ws_dm.cell(row=r, column=2).number_format = dollar_fmt
    ws_dm.cell(row=r, column=3).number_format = dollar_fmt
ws_dm['B37'].number_format = '0.00x'
ws_dm['C37'].number_format = '0.00x'
ws_dm['B43'].number_format = '0.0'
ws_dm['C43'].number_format = '0.0'

# --- Section 5: Verdict ---
ws_dm['A45'] = 'VERDICT'
style_header_row(ws_dm, 45, 6)
ws_dm['A46'] = 'EV Advantage (Path 3 - Path 2)'
ws_dm['B46'] = '=F31-F19'
ws_dm['B46'].number_format = dollar_fmt
ws_dm['B46'].font = Font(bold=True, size=13)

ws_dm['A47'] = 'EV Multiple (Path 3 / Path 2)'
ws_dm['B47'] = '=F31/F19'
ws_dm['B47'].number_format = '0.0x'
ws_dm['B47'].font = Font(bold=True, size=13)

ws_dm['A48'] = 'Recommended Path'
ws_dm['B48'] = '=IF(F31>F19,"PATH 3 — Roll into Evecxia","PATH 2 — Hold LBRX")'
ws_dm['B48'].font = Font(bold=True, size=13, color='006600')

for r in range(46, 49):
    ws_dm.cell(row=r, column=1).font = Font(bold=True, size=11)
    for c in range(1, 3):
        ws_dm.cell(row=r, column=c).border = thin_border

# --- Section 6: Probability Sensitivity ---
ws_dm['A50'] = 'SENSITIVITY: EV BY CLINICAL PROBABILITY'
style_header_row(ws_dm, 50, 8)

ws_dm['A51'] = 'Path 3 Net EV at varying P(Ph2 Success):'
ws_dm['A51'].font = Font(italic=True, size=10)

prob_values = [0.20, 0.25, 0.30, 0.35, 0.40, 0.45, 0.50]
for i, p in enumerate(prob_values):
    col = i + 2
    # Header: probability
    cell = ws_dm.cell(row=52, column=col, value=p)
    cell.number_format = pct_fmt
    cell.font = header_font
    cell.fill = light_green_fill
    cell.border = thin_border
    cell.alignment = Alignment(horizontal='center')

    # EV calculation: p * weighted_positive + (1-p) * weighted_negative
    # Positive weighted avg = 0.25*Low + 0.50*Base + 0.25*High
    # Negative weighted avg = 0.25*Low + 0.50*Base + 0.25*High
    col_letter = get_column_letter(col)
    ws_dm.cell(row=53, column=col,
        value=f'={col_letter}52*(0.25*$F$24/$B$24*{col_letter}52+0.50*$F$25/$B$25*{col_letter}52+0.25*$F$26/$B$26*{col_letter}52)+(1-{col_letter}52)*(0.25*$F$27/$B$27*(1-{col_letter}52)+0.50*$F$28/$B$28*(1-{col_letter}52)+0.25*$F$29/$B$29*(1-{col_letter}52))')

# Actually, simpler approach: just compute directly
for i, p in enumerate(prob_values):
    col = i + 2
    # Positive EV component (weighted across Low/Base/High within positive)
    # Negative EV component (weighted across Low/Base/High within negative)
    ws_dm.cell(row=53, column=col,
        value=f'={get_column_letter(col)}52*(0.25*$D$24+0.50*$D$25+0.25*$D$26)+(1-{get_column_letter(col)}52)*(0.25*$D$27+0.50*$D$28+0.25*$D$29)')
    ws_dm.cell(row=53, column=col).number_format = dollar_fmt
    ws_dm.cell(row=53, column=col).border = thin_border

ws_dm['A52'] = 'P(Ph2 Success)'
ws_dm['A52'].font = header_font
ws_dm['A52'].fill = light_green_fill
ws_dm['A52'].border = thin_border
ws_dm['A53'] = 'Path 3 Expected Net'
ws_dm['A53'].font = header_font
ws_dm['A53'].border = thin_border
ws_dm['A54'] = 'Path 2 Expected Net (for comparison)'
ws_dm['A54'].font = header_font
ws_dm['A54'].border = thin_border
for i in range(len(prob_values)):
    col = i + 2
    ws_dm.cell(row=54, column=col, value=f'=$F$19')
    ws_dm.cell(row=54, column=col).number_format = dollar_fmt
    ws_dm.cell(row=54, column=col).fill = light_blue_fill
    ws_dm.cell(row=54, column=col).border = thin_border

ws_dm['A55'] = 'Path 3 Favored?'
ws_dm['A55'].font = header_font
ws_dm['A55'].border = thin_border
for i in range(len(prob_values)):
    col = i + 2
    cl = get_column_letter(col)
    ws_dm.cell(row=55, column=col, value=f'=IF({cl}53>{cl}54,"YES","NO")')
    ws_dm.cell(row=55, column=col).border = thin_border
    ws_dm.cell(row=55, column=col).alignment = Alignment(horizontal='center')

print("   Decision Matrix added.")

# ============================================================
# 3. ADD BREAKEVEN ANALYSIS
# ============================================================
print("3. Adding Breakeven Analysis...")

if 'Breakeven' in wb.sheetnames:
    del wb['Breakeven']
ws_be = wb.create_sheet('Breakeven', wb.sheetnames.index('Outputs'))

ws_be.column_dimensions['A'].width = 42
ws_be.column_dimensions['B'].width = 22
ws_be.column_dimensions['C'].width = 22
ws_be.column_dimensions['D'].width = 22

ws_be['A1'] = 'BREAKEVEN ANALYSIS'
ws_be['A1'].font = title_font

# Q1: What Evecxia valuation matches Path 2 base case?
ws_be['A3'] = 'KEY QUESTION: AT WHAT EVECXIA VALUATION DOES PATH 3 = PATH 2?'
style_header_row(ws_be, 3, 3)

ws_be['A5'] = 'Path 2 Base Case Net (LBRX Ph3+ @ $100)'
ws_be['B5'] = '=Scenarios!C21'
ws_be['B5'].number_format = dollar_fmt

ws_be['A6'] = 'Investor Evecxia Shares'
ws_be['B6'] = '=Inputs!B32'
ws_be['B6'].number_format = number_fmt

ws_be['A7'] = 'Evecxia Total Shares at Readout'
ws_be['B7'] = '=Inputs!B65'
ws_be['B7'].number_format = number_fmt

ws_be['A8'] = 'Investor Ownership %'
ws_be['B8'] = '=B6/B7'
ws_be['B8'].number_format = '0.000%'

ws_be['A10'] = 'Breakeven Evecxia Valuation (to match Path 2 base)'
ws_be['B10'] = '=B5/B8'
ws_be['B10'].number_format = '$#,##0,,\"M\"'
ws_be['B10'].font = Font(bold=True, size=14)
ws_be['C10'] = '← Evecxia must be worth this much at Phase 2 readout to match LBRX Ph3+ base'
ws_be['C10'].font = Font(italic=True, color='666666', size=10)

ws_be['A12'] = 'Breakeven as Multiple of $180M IPO'
ws_be['B12'] = '=B10/Inputs!B64'
ws_be['B12'].number_format = '0.0x'
ws_be['B12'].font = Font(bold=True, size=12)

ws_be['A13'] = 'Breakeven as Multiple of $30M Series A'
ws_be['B13'] = '=B10/Inputs!B35'
ws_be['B13'].number_format = '0.0x'

for r in range(5, 14):
    for c in range(1, 3):
        ws_be.cell(row=r, column=c).border = thin_border

# Q2: What Evecxia valuation preserves original $100K?
ws_be['A15'] = 'CAPITAL PRESERVATION: WHAT VALUATION RETURNS YOUR $100K?'
style_header_row(ws_be, 15, 3)

ws_be['A17'] = 'Original Invested Capital'
ws_be['B17'] = '=serb_basis'
ws_be['B17'].number_format = dollar_fmt

ws_be['A18'] = 'Evecxia Valuation to Return $100K'
ws_be['B18'] = '=B17/B8'
ws_be['B18'].number_format = '$#,##0,,\"M\"'
ws_be['B18'].font = Font(bold=True, size=12)

ws_be['A19'] = 'As % of $180M IPO Valuation'
ws_be['B19'] = '=B18/Inputs!B64'
ws_be['B19'].number_format = pct_fmt

for r in range(17, 20):
    for c in range(1, 3):
        ws_be.cell(row=r, column=c).border = thin_border

# Q3: Sensitivity — Evecxia valuation vs investor net
ws_be['A21'] = 'SENSITIVITY: EVECXIA VALUATION → INVESTOR NET'
style_header_row(ws_be, 21, 4)

eve_vals = [50, 100, 150, 180, 250, 400, 600, 800, 1000, 1500, 2000, 3000]
ws_be['A22'] = 'Evecxia Valuation ($M)'
ws_be['B22'] = 'Investor Gross'
ws_be['C22'] = 'Tax'
ws_be['D22'] = 'Investor Net'
for c in range(1, 5):
    ws_be.cell(row=22, column=c).font = header_font
    ws_be.cell(row=22, column=c).fill = light_blue_fill
    ws_be.cell(row=22, column=c).border = thin_border

for i, val in enumerate(eve_vals):
    row = 23 + i
    ws_be.cell(row=row, column=1, value=val * 1000000)
    ws_be.cell(row=row, column=1).number_format = '$#,##0,,\"M\"'
    # Gross = valuation * ownership %
    ws_be.cell(row=row, column=2, value=f'=A{row}*$B$8')
    ws_be.cell(row=row, column=2).number_format = dollar_fmt
    # Tax = $0 (QSBS)
    ws_be.cell(row=row, column=3, value=0)
    ws_be.cell(row=row, column=3).number_format = dollar_fmt
    # Net
    ws_be.cell(row=row, column=4, value=f'=B{row}-C{row}')
    ws_be.cell(row=row, column=4).number_format = dollar_fmt
    for c in range(1, 5):
        ws_be.cell(row=row, column=c).border = thin_border

    # Highlight the breakeven row
    if val == 180:
        for c in range(1, 5):
            ws_be.cell(row=row, column=c).fill = light_yellow_fill

# Add MOIC column
ws_be['E22'] = 'MOIC'
ws_be['E22'].font = header_font
ws_be['E22'].fill = light_blue_fill
ws_be['E22'].border = thin_border
for i in range(len(eve_vals)):
    row = 23 + i
    ws_be.cell(row=row, column=5, value=f'=D{row}/serb_basis')
    ws_be.cell(row=row, column=5).number_format = '0.0x'
    ws_be.cell(row=row, column=5).border = thin_border

print("   Breakeven Analysis added.")

# ============================================================
# 4. ADD AFTER-TAX IRR
# ============================================================
print("4. Adding IRR calculations to Outputs...")

ws_out = wb['Outputs']

# Find the next available row after existing content
irr_start = 26

ws_out.cell(row=irr_start, column=1, value='AFTER-TAX IRR COMPARISON').font = title_font
style_header_row(ws_out, irr_start + 1, 7)

headers_irr = ['', 'S1: Lock-Up', 'S2: Ph3+', 'S3: Ph3-', 'S4: Dual', 'S5: Eve Ph2+', 'S6: Eve Ph2-']
for i, h in enumerate(headers_irr):
    ws_out.cell(row=irr_start + 2, column=i + 1, value=h).font = header_font
    ws_out.cell(row=irr_start + 2, column=i + 1).fill = light_blue_fill
    ws_out.cell(row=irr_start + 2, column=i + 1).border = thin_border

r = irr_start + 3
ws_out.cell(row=r, column=1, value='Investment Date').font = header_font
ws_out.cell(row=r, column=1).border = thin_border
for col in range(2, 8):
    ws_out.cell(row=r, column=col, value='=serb_date')
    ws_out.cell(row=r, column=col).number_format = 'YYYY-MM-DD'
    ws_out.cell(row=r, column=col).border = thin_border

r = irr_start + 4
ws_out.cell(row=r, column=1, value='Exit Date').font = header_font
ws_out.cell(row=r, column=1).border = thin_border
exit_formulas = ['=lockup_date', '=ph3_date', '=ph3_date', '=lockup_date', '=evecxia_ph2_date', '=evecxia_ph2_date']
for i, f in enumerate(exit_formulas):
    ws_out.cell(row=r, column=i + 2, value=f)
    ws_out.cell(row=r, column=i + 2).number_format = 'YYYY-MM-DD'
    ws_out.cell(row=r, column=i + 2).border = thin_border

r = irr_start + 5
ws_out.cell(row=r, column=1, value='Holding Period (years)').font = header_font
ws_out.cell(row=r, column=1).border = thin_border
for col in range(2, 8):
    cl = get_column_letter(col)
    ws_out.cell(row=r, column=col, value=f'=({cl}{irr_start+4}-{cl}{irr_start+3})/365.25')
    ws_out.cell(row=r, column=col).number_format = '0.00'
    ws_out.cell(row=r, column=col).border = thin_border

r = irr_start + 6
ws_out.cell(row=r, column=1, value='Cost Basis').font = header_font
ws_out.cell(row=r, column=1).border = thin_border
for col in range(2, 8):
    ws_out.cell(row=r, column=col, value='=serb_basis')
    ws_out.cell(row=r, column=col).number_format = dollar_fmt
    ws_out.cell(row=r, column=col).border = thin_border

r = irr_start + 7
ws_out.cell(row=r, column=1, value='Net After-Tax').font = header_font
ws_out.cell(row=r, column=1).border = thin_border
# Pull from Scenarios
scenario_cols = ['B', 'C', 'D', 'E', 'F', 'G']
for i, sc in enumerate(scenario_cols):
    ws_out.cell(row=r, column=i + 2, value=f'=Scenarios!{sc}21')
    ws_out.cell(row=r, column=i + 2).number_format = dollar_fmt
    ws_out.cell(row=r, column=i + 2).border = thin_border

r = irr_start + 8
ws_out.cell(row=r, column=1, value='After-Tax IRR (annualized)').font = Font(bold=True, size=12)
ws_out.cell(row=r, column=1).border = thin_border
for col in range(2, 8):
    cl = get_column_letter(col)
    # IRR = (ending/beginning)^(1/years) - 1
    ws_out.cell(row=r, column=col,
        value=f'=IF({cl}{irr_start+7}>0,({cl}{irr_start+7}/{cl}{irr_start+6})^(1/{cl}{irr_start+5})-1,-1)')
    ws_out.cell(row=r, column=col).number_format = '0.0%'
    ws_out.cell(row=r, column=col).font = Font(bold=True, size=12)
    ws_out.cell(row=r, column=col).border = thin_border

r = irr_start + 10
ws_out.cell(row=r, column=1, value='Note: IRR assumes single lump-sum in/out. Negative returns shown as -100%.').font = Font(italic=True, color='666666', size=9)

print("   IRR calculations added.")

# ============================================================
# 5. ADD DILUTION WATERFALL
# ============================================================
print("5. Adding Dilution Waterfall...")

if 'Dilution' in wb.sheetnames:
    del wb['Dilution']
ws_dil = wb.create_sheet('Dilution', wb.sheetnames.index('Outputs'))

ws_dil.column_dimensions['A'].width = 32
for c in ['B','C','D','E','F']:
    ws_dil.column_dimensions[c].width = 20

ws_dil['A1'] = 'EVECXIA DILUTION WATERFALL'
ws_dil['A1'].font = title_font

ws_dil['A2'] = 'Tracking investor ownership from Series A entry through Phase 2 readout'
ws_dil['A2'].font = Font(italic=True, color='666666', size=10)

headers_dil = ['Stage', 'Total Shares', 'Investor Shares', 'Ownership %', 'Implied Value of Stake']
for i, h in enumerate(headers_dil, 1):
    ws_dil.cell(row=4, column=i, value=h).font = header_font
    ws_dil.cell(row=4, column=i).fill = section_fill
    ws_dil.cell(row=4, column=i).font = section_font
    ws_dil.cell(row=4, column=i).border = thin_border
    ws_dil.cell(row=4, column=i).alignment = Alignment(horizontal='center')

# Stage 1: Series A Close
ws_dil['A5'] = 'Series A Close ($30M post-money)'
ws_dil['B5'] = '=Inputs!B36'  # Total shares at SerA close
ws_dil['C5'] = '=Inputs!B32'  # Investor shares
ws_dil['D5'] = '=C5/B5'
ws_dil['E5'] = '=D5*Inputs!B35'  # ownership * post-money

# Stage 2: Bridge Round (assume 10% dilution)
ws_dil['A6'] = 'Bridge Round (~10% dilution, est.)'
ws_dil['B6'] = '=B5/0.9'  # shares increase to dilute 10%
ws_dil['C6'] = '=C5'  # investor shares unchanged
ws_dil['D6'] = '=C6/B6'
ws_dil['E6'] = '=D6*45000000'  # assume $45M post-bridge

# Stage 3: IPO ($180M)
ws_dil['A7'] = 'IPO ($180M valuation)'
ws_dil['B7'] = '=B6/0.80'  # IPO creates ~20% new float
ws_dil['C7'] = '=C5'
ws_dil['D7'] = '=C7/B7'
ws_dil['E7'] = '=D7*Inputs!B64'

# Stage 4: At Readout (365M shares est.)
ws_dil['A8'] = 'At Phase 2 Readout (365M shares est.)'
ws_dil['B8'] = '=Inputs!B65'
ws_dil['C8'] = '=C5'
ws_dil['D8'] = '=C8/B8'
ws_dil['E8'] = 'varies by scenario'

for r in range(5, 9):
    ws_dil.cell(row=r, column=2).number_format = '#,##0'
    ws_dil.cell(row=r, column=3).number_format = '#,##0'
    ws_dil.cell(row=r, column=4).number_format = '0.000%'
    ws_dil.cell(row=r, column=5).number_format = dollar_fmt
    for c in range(1, 6):
        ws_dil.cell(row=r, column=c).border = thin_border

# Highlight key insight
ws_dil['A10'] = 'KEY INSIGHT'
ws_dil['A10'].font = Font(bold=True, size=12)
ws_dil['A11'] = 'Ownership dilutes from Series A to readout:'
ws_dil['A12'] = 'Series A ownership'
ws_dil['B12'] = '=D5'
ws_dil['B12'].number_format = '0.000%'
ws_dil['A13'] = 'Readout ownership'
ws_dil['B13'] = '=D8'
ws_dil['B13'].number_format = '0.000%'
ws_dil['A14'] = 'Dilution (% of original ownership lost)'
ws_dil['B14'] = '=1-D8/D5'
ws_dil['B14'].number_format = pct_fmt
ws_dil['B14'].font = Font(bold=True, size=12, color='CC0000')

for r in range(12, 15):
    for c in range(1, 3):
        ws_dil.cell(row=r, column=c).border = thin_border

# Readout value at each scenario
ws_dil['A16'] = 'STAKE VALUE AT READOUT BY SCENARIO'
style_header_row(ws_dil, 16, 4)

scenario_names = ['Ph2+ Low ($800M)', 'Ph2+ Base ($1.5B)', 'Ph2+ High ($3B)',
                  'Ph2- Low ($50M)', 'Ph2- Base ($120M)', 'Ph2- High ($250M)']
val_refs = ['Inputs!B58', 'Inputs!B59', 'Inputs!B60', 'Inputs!B61', 'Inputs!B62', 'Inputs!B63']
fills = [light_green_fill]*3 + [light_red_fill]*3

ws_dil['A17'] = 'Scenario'
ws_dil['B17'] = 'Company Valuation'
ws_dil['C17'] = 'Investor Stake Value'
ws_dil['D17'] = 'MOIC'
for c in range(1, 5):
    ws_dil.cell(row=17, column=c).font = header_font
    ws_dil.cell(row=17, column=c).fill = light_blue_fill
    ws_dil.cell(row=17, column=c).border = thin_border

for i, (name, ref, fill) in enumerate(zip(scenario_names, val_refs, fills)):
    row = 18 + i
    ws_dil.cell(row=row, column=1, value=name)
    ws_dil.cell(row=row, column=2, value=f'={ref}')
    ws_dil.cell(row=row, column=2).number_format = '$#,##0,,\"M\"'
    ws_dil.cell(row=row, column=3, value=f'=B{row}*$D$8')
    ws_dil.cell(row=row, column=3).number_format = dollar_fmt
    ws_dil.cell(row=row, column=4, value=f'=C{row}/serb_basis')
    ws_dil.cell(row=row, column=4).number_format = '0.0x'
    for c in range(1, 5):
        ws_dil.cell(row=row, column=c).border = thin_border
        ws_dil.cell(row=row, column=c).fill = fill

print("   Dilution Waterfall added.")

# ============================================================
# 6. ENHANCE SENSITIVITY IN OUTPUTS
# ============================================================
print("6. Enhancing Sensitivity tables...")

ws_out = wb['Outputs']

# Add 2D sensitivity: P(success) vs Evecxia valuation multiplier
sens_start = irr_start + 12

ws_out.cell(row=sens_start, column=1, value='2D SENSITIVITY: P(Ph2+) × EVECXIA POST-DATA VALUATION').font = title_font

# Column headers: Evecxia valuations
eve_val_labels = ['$400M', '$600M', '$800M', '$1.0B', '$1.5B', '$2.0B', '$3.0B']
eve_val_nums = [400000000, 600000000, 800000000, 1000000000, 1500000000, 2000000000, 3000000000]

ws_out.cell(row=sens_start + 1, column=1, value='Investor Net After-Tax ($)').font = Font(italic=True, size=10)
ws_out.cell(row=sens_start + 2, column=1, value='P(Ph2+) ↓  /  Valuation →').font = header_font
ws_out.cell(row=sens_start + 2, column=1).fill = light_yellow_fill
ws_out.cell(row=sens_start + 2, column=1).border = thin_border

for j, (label, val) in enumerate(zip(eve_val_labels, eve_val_nums)):
    col = j + 2
    cell = ws_out.cell(row=sens_start + 2, column=col, value=val)
    cell.number_format = '$#,##0,,\"M\"'
    cell.font = header_font
    cell.fill = light_yellow_fill
    cell.border = thin_border
    cell.alignment = Alignment(horizontal='center')

# Row headers: probabilities
prob_rows = [0.20, 0.25, 0.30, 0.35, 0.40, 0.45, 0.50]
for i, p in enumerate(prob_rows):
    row = sens_start + 3 + i
    cell = ws_out.cell(row=row, column=1, value=p)
    cell.number_format = pct_fmt
    cell.font = header_font
    cell.fill = light_yellow_fill
    cell.border = thin_border

    for j, val in enumerate(eve_val_nums):
        col = j + 2
        # Net = p * (valuation * ownership%) + (1-p) * (neg_base_val * ownership%)
        # ownership% = Inputs!B32 / Inputs!B65
        # neg_base_val = Inputs!B62 = $120M
        cell = ws_out.cell(row=row, column=col,
            value=f'=$A{row}*({get_column_letter(col)}${sens_start+2}*Inputs!$B$32/Inputs!$B$65)+(1-$A{row})*(Inputs!$B$62*Inputs!$B$32/Inputs!$B$65)')
        cell.number_format = dollar_fmt
        cell.border = thin_border
        # Color code: green if > Path 2 base, red if < basis
        # (Can't do conditional formatting via openpyxl easily, so leave for Excel)

print("   Enhanced Sensitivity added.")

# ============================================================
# SAVE
# ============================================================
wb.save('/home/user/LB-Pharma-qsbs/LBRX_QSBS_Evecxia_Model.xlsx')
print("\nAll enhancements saved successfully.")
