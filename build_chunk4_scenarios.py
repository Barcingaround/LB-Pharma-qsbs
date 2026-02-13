#!/usr/bin/env python3
"""
LBRX QSBS Evecxia Model - Chunk 4: Scenarios Tab
Builds 6 scenarios (S1-S6) with full tax calculations
S1: Series B sells at lock-up (flat)
S2: Series B holds through Phase 3 positive
S3: Series B holds through Phase 3 negative
S4: Series B + Series C sells at lock-up (lot-by-lot)
S5: §1045 rollover → Evecxia Phase 2 positive
S6: §1045 rollover → Evecxia Phase 2 negative
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime

def build_scenarios_tab(wb):
    """Build Scenarios tab with 6 scenarios"""

    ws = wb['Scenarios']

    # Set column widths
    ws.column_dimensions['A'].width = 35
    for col_letter in ['B', 'C', 'D', 'E', 'F', 'G']:
        ws.column_dimensions[col_letter].width = 18

    def set_header(row, text):
        ws[f'A{row}'] = text
        ws[f'A{row}'].font = Font(name='Calibri', size=12, bold=True, color='FFFFFF')
        ws[f'A{row}'].fill = PatternFill(start_color='2E75B6', end_color='2E75B6', fill_type='solid')
        ws.merge_cells(f'A{row}:G{row}')
        return row + 1

    def set_label(row, text, indent=0):
        ws[f'A{row}'] = '  ' * indent + text
        ws[f'A{row}'].font = Font(name='Calibri', size=10)
        return row

    def set_value(row, col, formula, number_format=None):
        cell = ws.cell(row=row, column=col)
        cell.value = formula
        if number_format:
            cell.number_format = number_format
        cell.alignment = Alignment(horizontal='right', vertical='center')

    row = 1

    # Title
    ws['A1'] = 'SCENARIO ANALYSIS'
    ws['A1'].font = Font(name='Calibri', size=14, bold=True)
    ws.merge_cells('A1:G1')
    row += 2

    # Column headers for scenario comparison
    headers_row = row
    scenario_headers = ['Metric', 'S1: Lock-Up Sale', 'S2: Ph3+ Hold', 'S3: Ph3- Hold', 'S4: Dual Lot Lock-Up', 'S5: §1045→Eve Ph2+', 'S6: §1045→Eve Ph2-']
    for col_num, header in enumerate(scenario_headers, 1):
        cell = ws.cell(row=headers_row, column=col_num)
        cell.value = header
        cell.font = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    row += 1

    # Metrics
    metrics = [
        ('Sale Date', ['=lockup_date', '=ph3_date', '=ph3_date', '=lockup_date', '=evecxia_ph2_date', '=evecxia_ph2_date'], 'yyyy-mm-dd'),
        ('Sale Price/Share', ['=lockup_price', '=Inputs!$B$53', '=Inputs!$B$56', '=lockup_price', 'N/A (Evecxia)', 'N/A (Evecxia)'], '$#,##0.00'),
        ('Shares Sold', ['=serb_shares', '=serb_shares', '=serb_shares', '=serb_shares+B25', 'Evecxia', 'Evecxia'], '#,##0'),
        ('Gross Proceeds', ['=serb_shares*lockup_price', '=serb_shares*Inputs!$B$53', '=serb_shares*Inputs!$B$56', '=B9+C9', '=Inputs!$B$65*0.0026*Inputs!$B$59', '=Inputs!$B$65*0.0026*Inputs!$B$62'], '$#,##0'),
        ('Cost Basis', ['=serb_basis', '=serb_basis', '=serb_basis', '=serb_basis+C25', '=serb_basis', '=serb_basis'], '$#,##0'),
        ('Gross Gain/(Loss)', ['=B9-B10', '=C9-C10', '=D9-D10', '=E9-E10', '=F9-F10', '=G9-G10'], '$#,##0'),
        ('', ['', '', '', '', '', ''], None),  # Blank row
        ('Holding Period (days)', ['=lockup_date-serb_date', '=ph3_date-serb_date', '=ph3_date-serb_date', 'Varies', '=evecxia_ph2_date-serb_date', '=evecxia_ph2_date-serb_date'], '#,##0'),
        ('Holding Period (years)', ['=B13/365.25', '=C13/365.25', '=D13/365.25', 'Varies', '=F13/365.25', '=G13/365.25'], '0.00'),
        ('Long-Term?', ['=IF(B13>365,"YES","NO")', '=IF(C13>365,"YES","NO")', '=IF(D13>365,"YES","NO")', 'YES', 'YES', 'YES'], None),
        ('QSBS 5-Year Met?', ['=IF(B13>=1826,"YES","NO")', '=IF(C13>=1826,"YES","NO")', '=IF(D13>=1826,"YES","NO")', 'See detail', '=IF(F13>=1826,"YES","NO")', '=IF(G13>=1826,"YES","NO")'], None),
        ('', ['', '', '', '', '', ''], None),  # Blank row
        ('§1202 Exclusion Amount', ['=IF(B16="YES",MIN(B11,10000000),0)', '=IF(C16="YES",MIN(C11,10000000),0)', '=IF(D16="YES",MAX(0,MIN(D11,10000000)),0)', 'See detail', '=IF(F16="YES",MIN(F11,10000000),0)', '=IF(G16="YES",MIN(MAX(0,G11),10000000),0)'], '$#,##0'),
        ('§1045 Deferred Gain', ['0', '0', '0', '0', '=B11', '=B11'], '$#,##0'),
        ('Taxable Gain', ['=MAX(0,B11-B18-B19)', '=MAX(0,C11-C18-C19)', '=MAX(0,D11-D18-D19)', 'See detail', '=MAX(0,F11-F18-F19)', '=MAX(0,G11-G18-G19)'], '$#,##0'),
        ('', ['', '', '', '', '', ''], None),  # Blank row
        ('Federal LTCG Tax', ['=B20*eff_fed_rate', '=C20*eff_fed_rate', '=D20*eff_fed_rate', 'See detail', '=F20*eff_fed_rate', '=G20*eff_fed_rate'], '$#,##0'),
        ('State Tax', ['=B20*state_rate', '=C20*state_rate', '=D20*state_rate', 'See detail', '=F20*state_rate', '=G20*state_rate'], '$#,##0'),
        ('Total Tax', ['=B22+B23', '=C22+C23', '=D22+D23', 'See detail', '=F22+F23', '=G22+G23'], '$#,##0'),
        ('', ['', '', '', '', '', ''], None),  # Blank row
        ('Net After-Tax Proceeds', ['=B9-B24', '=C9-C24', '=D9-D24', 'See detail', '=F9-F24', '=G9-G24'], '$#,##0'),
        ('MOIC (Multiple on Investment)', ['=B26/serb_basis', '=C26/serb_basis', '=D26/serb_basis', 'Varies', '=F26/serb_basis', '=G26/serb_basis'], '0.00x'),
        ('Effective Tax Rate', ['=IF(B11>0,B24/B11,0)', '=IF(C11>0,C24/C11,0)', '=IF(D11>0,D24/D11,0)', 'Varies', '=IF(F11>0,F24/F11,0)', '=IF(G11>0,G24/G11,0)'], '0.0%'),
    ]

    for metric_label, formulas, num_fmt in metrics:
        set_label(row, metric_label)
        for col_num, formula in enumerate(formulas, 2):
            set_value(row, col_num, formula, num_fmt)
        row += 1

    row += 2

    # S4 DETAIL: Dual lot (Series B + Series C)
    row = set_header(row, 'SCENARIO S4 DETAIL: Series B + Series C Lot-by-Lot at Lock-Up')
    set_label(row, 'Metric'); ws['B' + str(row)] = 'Series B Lot'; ws['C' + str(row)] = 'Series C Lot'; ws['D' + str(row)] = 'Total'
    for col in ['A', 'B', 'C', 'D']:
        ws[col + str(row)].font = Font(name='Calibri', size=10, bold=True)
    row += 1

    s4_metrics = [
        ('Sale Date', ['=lockup_date', '=lockup_date', 'Same'], None),
        ('Sale Price/Share', ['=lockup_price', '=lockup_price', 'Same'], '$#,##0.00'),
        ('Shares Sold', ['=serb_shares', '=IF(Inputs!$B$23="","",Inputs!$B$25)', '=B{r}+C{r}'], '#,##0'),
        ('Gross Proceeds', ['=B{r1}*lockup_price', '=C{r1}*lockup_price', '=B{r}+C{r}'], '$#,##0'),
        ('Cost Basis', ['=serb_basis', '=IF(Inputs!$B$23="","",Inputs!$B$23)', '=B{r}+C{r}'], '$#,##0'),
        ('Gross Gain', ['=B{r1}-B{r2}', '=C{r1}-C{r2}', '=B{r}+C{r}'], '$#,##0'),
        ('Holding Period (days)', ['=lockup_date-serb_date', '=IF(Inputs!$B$22="","",lockup_date-Inputs!$B$22)', 'Varies'], '#,##0'),
        ('QSBS Eligible?', ['YES', 'NO', 'Mixed'], None),
        ('QSBS 5-Year Met?', ['=IF(B{r1}>=1826,"YES","NO")', 'N/A', 'Mixed'], None),
        ('§1202 Exclusion', ['=IF(B{r1}="YES",MIN(B{r3},10000000),0)', '0', '=B{r}+C{r}'], '$#,##0'),
        ('Taxable Gain', ['=MAX(0,B{r4}-B{r1})', '=MAX(0,C{r4})', '=B{r}+C{r}'], '$#,##0'),
        ('Tax Due', ['=B{r1}*combined_rate', '=C{r1}*combined_rate', '=B{r}+C{r}'], '$#,##0'),
        ('Net After-Tax', ['=B{r8}-B{r1}', '=C{r8}-C{r1}', '=B{r}+C{r}'], '$#,##0'),
    ]

    start_row = row
    for i, (label, formulas, num_fmt) in enumerate(s4_metrics):
        set_label(row, label)
        for col_num, formula in enumerate(formulas, 2):
            if '{r}' in formula:
                # Replace {r} with current row, {r1} with row-1, etc.
                formula = formula.replace('{r}', str(row)).replace('{r1}', str(row-1)).replace('{r2}', str(row-2)).replace('{r3}', str(row-3)).replace('{r4}', str(row-4)).replace('{r8}', str(row-8))
            set_value(row, col_num, formula, num_fmt)
        row += 1

    row += 2

    # S5/S6 DETAIL: Two-leg journey
    row = set_header(row, 'SCENARIOS S5/S6 DETAIL: §1045 Rollover → Evecxia (Two-Leg Journey)')
    set_label(row, 'Leg / Metric'); ws['B' + str(row)] = 'S5: Positive Ph2'; ws['C' + str(row)] = 'S6: Negative Ph2'
    for col in ['A', 'B', 'C']:
        ws[col + str(row)].font = Font(name='Calibri', size=10, bold=True)
    row += 1

    set_label(row, 'LEG 1: LBRX SALE AT LOCK-UP', 1)
    ws['A' + str(row)].font = Font(name='Calibri', size=10, bold=True)
    row += 1

    leg1_metrics = [
        ('LBRX Sale Date', ['=lockup_date', '=lockup_date'], 'yyyy-mm-dd'),
        ('LBRX Sale Price', ['=lockup_price', '=lockup_price'], '$#,##0.00'),
        ('LBRX Gross Proceeds', ['=serb_shares*lockup_price', '=serb_shares*lockup_price'], '$#,##0'),
        ('LBRX Cost Basis', ['=serb_basis', '=serb_basis'], '$#,##0'),
        ('LBRX Realized Gain', ['=B{r1}-B{r2}', '=C{r1}-C{r2}'], '$#,##0'),
        ('§1045 Deferred Gain', ['=B{r1}', '=C{r1}'], '$#,##0'),
        ('Tax at LBRX Sale', ['0', '0'], '$#,##0'),
    ]

    for label, formulas, num_fmt in leg1_metrics:
        set_label(row, label, 2)
        for col_num, formula in enumerate(formulas, 2):
            if '{r1}' in formula:
                formula = formula.replace('{r1}', str(row-1)).replace('{r2}', str(row-2))
            set_value(row, col_num, formula, num_fmt)
        row += 1

    row += 1
    set_label(row, 'LEG 2: EVECXIA PURCHASE & SALE', 1)
    ws['A' + str(row)].font = Font(name='Calibri', size=10, bold=True)
    row += 1

    leg2_metrics = [
        ('Evecxia Purchase Date', ['=evecxia_purchase', '=evecxia_purchase'], 'yyyy-mm-dd'),
        ('Amount Reinvested', ['=serb_basis', '=serb_basis'], '$#,##0'),
        ('Evecxia Share Price at Purchase', ['=Inputs!$B$30', '=Inputs!$B$30'], '$#,##0.0000'),
        ('Evecxia Shares Acquired', ['=B{r1}/B{r2}', '=C{r1}/C{r2}'], '#,##0'),
        ('Evecxia Adjusted Basis (§1045)', ['=B{r2}-B{r14}', '=C{r2}-C{r14}'], '$#,##0'),
        ('', ['', ''], None),
        ('Evecxia Sale Date', ['=evecxia_ph2_date', '=evecxia_ph2_date'], 'yyyy-mm-dd'),
        ('Evecxia Valuation at Sale', ['=Inputs!$B$59', '=Inputs!$B$62'], '$#,##0'),
        ('Evecxia Share Price at Sale', ['=B{r1}/Inputs!$B$65', '=C{r1}/Inputs!$B$65'], '$#,##0.00'),
        ('Evecxia Gross Proceeds', ['=B{r6}*B{r1}', '=C{r6}*C{r1}'], '$#,##0'),
        ('Evecxia Gain', ['=B{r1}-B{r7}', '=C{r1}-C{r7}'], '$#,##0'),
        ('Tacked Holding (days)', ['=evecxia_ph2_date-serb_date', '=evecxia_ph2_date-serb_date'], '#,##0'),
        ('QSBS 5-Year Met?', ['=IF(B{r1}>=1826,"YES","NO")', '=IF(C{r1}>=1826,"YES","NO")'], None),
        ('§1202 Exclusion', ['=IF(B{r1}="YES",MIN(B{r3},10000000),0)', '=IF(C{r1}="YES",MIN(MAX(0,C{r3}),10000000),0)'], '$#,##0'),
        ('Taxable Gain on Evecxia', ['=MAX(0,B{r5}-B{r1})', '=MAX(0,C{r5}-C{r1})'], '$#,##0'),
        ('Tax on Evecxia Sale', ['=B{r1}*combined_rate', '=C{r1}*combined_rate'], '$#,##0'),
    ]

    for label, formulas, num_fmt in leg2_metrics:
        set_label(row, label, 2)
        for col_num, formula in enumerate(formulas, 2):
            if '{r' in formula:
                formula = (formula.replace('{r1}', str(row-1)).replace('{r2}', str(row-2))
                          .replace('{r3}', str(row-3)).replace('{r5}', str(row-5))
                          .replace('{r6}', str(row-6)).replace('{r7}', str(row-7))
                          .replace('{r14}', str(row-14)))
            set_value(row, col_num, formula, num_fmt)
        row += 1

    row += 1
    set_label(row, 'COMBINED JOURNEY TOTAL', 1)
    ws['A' + str(row)].font = Font(name='Calibri', size=11, bold=True)
    row += 1

    combined_metrics = [
        ('Original Investment', ['=serb_basis', '=serb_basis'], '$#,##0'),
        ('Final Proceeds', ['=B{r15}', '=C{r15}'], '$#,##0'),
        ('Total Tax (both legs)', ['=B{r17}+B{r1}', '=C{r17}+C{r1}'], '$#,##0'),
        ('Net After-Tax', ['=B{r2}-B{r1}', '=C{r2}-C{r1}'], '$#,##0'),
        ('MOIC', ['=B{r1}/serb_basis', '=C{r1}/serb_basis'], '0.00x'),
        ('Effective Tax Rate', ['=B{r3}/B{r2}', '=C{r3}/C{r2}'], '0.0%'),
    ]

    for label, formulas, num_fmt in combined_metrics:
        set_label(row, label, 2)
        for col_num, formula in enumerate(formulas, 2):
            if '{r' in formula:
                formula = (formula.replace('{r1}', str(row-1)).replace('{r2}', str(row-2))
                          .replace('{r3}', str(row-3)).replace('{r15}', str(row-15)).replace('{r17}', str(row-17)))
            set_value(row, col_num, formula, num_fmt)
        row += 1

    # Freeze panes
    ws.freeze_panes = 'A4'


def main():
    print("Building LBRX QSBS Evecxia Model - Chunk 4: Scenarios Tab...")

    # Load checkpoint
    wb = load_workbook('/home/user/LB-Pharma-qsbs/LBRX_QSBS_Evecxia_Model_v3_timeline_comps.xlsx')
    print("✓ Loaded checkpoint v3")

    # Build Scenarios tab
    build_scenarios_tab(wb)
    print("✓ Built Scenarios tab (6 scenarios: S1-S6 with full lot-by-lot calcs)")

    # Save checkpoint
    output_file = '/home/user/LB-Pharma-qsbs/LBRX_QSBS_Evecxia_Model_v4_scenarios.xlsx'
    wb.save(output_file)
    print(f"✓ Saved checkpoint: {output_file}")
    print("\nChunk 4 complete!")

if __name__ == '__main__':
    main()
