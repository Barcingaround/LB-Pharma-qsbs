#!/usr/bin/env python3
"""
LBRX QSBS Evecxia Model - Chunk 2: TaxLogic Tab
Implements 4 sections: Standard cap gains, §1202 checker, §1045 rollover, Combined chain
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime

def build_taxlogic_tab(wb):
    """Build TaxLogic tab with 4 sections of tax calculation logic"""

    ws = wb['TaxLogic']

    # Set column widths
    ws.column_dimensions['A'].width = 45
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 20

    # Helper functions
    def set_header(row, text, col='A'):
        cell = ws[f'{col}{row}']
        cell.value = text
        cell.font = Font(name='Calibri', size=12, bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='2E75B6', end_color='2E75B6', fill_type='solid')
        cell.alignment = Alignment(horizontal='left', vertical='center')

    def set_label(row, text):
        cell = ws[f'A{row}']
        cell.value = text
        cell.font = Font(name='Calibri', size=11)
        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

    def set_formula(row, formula, number_format=None):
        cell = ws[f'B{row}']
        cell.value = formula
        cell.alignment = Alignment(horizontal='right', vertical='center')
        if number_format:
            cell.number_format = number_format

    def set_result_label(row, text):
        cell = ws[f'C{row}']
        cell.value = text
        cell.font = Font(name='Calibri', size=10, italic=True)
        cell.alignment = Alignment(horizontal='left', vertical='center')

    row = 1

    # Title
    ws[f'A{row}'] = 'TAX LOGIC ENGINE'
    ws[f'A{row}'].font = Font(name='Calibri', size=14, bold=True)
    row += 1

    ws[f'A{row}'] = 'All tax calculations reference Inputs tab via named ranges'
    ws[f'A{row}'].font = Font(name='Calibri', size=10, italic=True)
    row += 2

    # ============================================================
    # SECTION 1: STANDARD CAPITAL GAINS TAX CALCULATOR
    # ============================================================
    set_header(row, 'SECTION 1: STANDARD CAPITAL GAINS TAX CALCULATOR')
    row += 1

    set_label(row, 'Example Inputs (for illustration):')
    row += 1
    set_label(row, '  Proceeds')
    set_formula(row, 100000, '$#,##0')
    row += 1
    set_label(row, '  Cost Basis')
    set_formula(row, 80000, '$#,##0')
    row += 1
    set_label(row, '  Holding Period (days)')
    set_formula(row, 400, '#,##0')
    row += 2

    set_label(row, 'CALCULATED:')
    row += 1
    set_label(row, 'Gain / (Loss)')
    set_formula(row, '=B6-B8', '$#,##0')
    row += 1
    set_label(row, 'Long-Term (>365 days)?')
    set_formula(row, '=IF(B10>365,"YES","NO")')
    row += 1
    set_label(row, 'Tax Rate Applied')
    set_formula(row, '=IF(B15="YES",combined_rate,Inputs!$B$11)', '0.0%')
    set_result_label(row, '← LTCG if >1yr, else ordinary')
    row += 1
    set_label(row, 'Tax Due')
    set_formula(row, '=MAX(0,B13)*B16', '$#,##0')
    row += 2

    # ============================================================
    # SECTION 2: §1202 QSBS ELIGIBILITY CHECK
    # ============================================================
    set_header(row, 'SECTION 2: §1202 QSBS ELIGIBILITY CHECK')
    row += 1

    set_label(row, 'Inputs (example):')
    row += 1
    set_label(row, '  Stock QSBS Eligible?')
    set_formula(row, 'YES')
    row += 1
    set_label(row, '  Holding Period (days)')
    set_formula(row, 2000, '#,##0')
    row += 1
    set_label(row, '  Gain Amount')
    set_formula(row, 500000, '$#,##0')
    row += 1
    set_label(row, '  Original Basis')
    set_formula(row, 100000, '$#,##0')
    row += 2

    set_label(row, 'ELIGIBILITY CHECKS:')
    row += 1
    set_label(row, 'Check A: Stock is QSBS?')
    set_formula(row, '=IF(B23="YES","PASS","FAIL")')
    set_result_label(row, '← Must be QSBS')
    row += 1
    set_label(row, 'Check B: Held ≥ 5 years (1826 days)?')
    set_formula(row, '=IF(B25>=1826,"PASS","FAIL")')
    set_result_label(row, '← 5 years = 1826 days')
    row += 1
    set_label(row, 'Overall Eligible?')
    set_formula(row, '=IF(AND(B32="PASS",B34="PASS"),"YES - §1202 AVAILABLE","NO")')
    row += 2

    set_label(row, '§1202 EXCLUSION CALCULATION:')
    row += 1
    set_label(row, 'Exclusion Cap')
    set_formula(row, '=MAX(10000000,10*B28)', '$#,##0')
    set_result_label(row, '← Greater of $10M or 10× basis')
    row += 1
    set_label(row, 'Gain to Exclude')
    set_formula(row, '=IF(B36="YES - §1202 AVAILABLE",MIN(B27,B40),0)', '$#,##0')
    row += 1
    set_label(row, 'Remaining Taxable Gain')
    set_formula(row, '=MAX(0,B27-B42)', '$#,##0')
    row += 1
    set_label(row, 'Tax on Remaining Gain')
    set_formula(row, '=B44*combined_rate', '$#,##0')
    row += 2

    # ============================================================
    # SECTION 3: §1045 ROLLOVER CALCULATOR (8-STEP)
    # ============================================================
    set_header(row, 'SECTION 3: §1045 ROLLOVER CALCULATOR (8-STEP WALKTHROUGH)')
    row += 1

    set_label(row, 'STEP 1: DETERMINE §1045 ELIGIBILITY')
    row += 1
    set_label(row, 'Inputs (example):')
    row += 1
    set_label(row, '  Original stock QSBS?')
    set_formula(row, 'YES')
    row += 1
    set_label(row, '  Held original > 6 months (180 days)?')
    set_formula(row, 'YES')
    row += 1
    set_label(row, '  Replacement purchased within 60 days?')
    set_formula(row, 'YES')
    row += 1
    set_label(row, '  Replacement stock is QSBS?')
    set_formula(row, 'YES')
    row += 2

    set_label(row, 'All §1045 Requirements Met?')
    set_formula(row, '=IF(AND(B53="YES",B55="YES",B57="YES",B59="YES"),"YES - §1045 ELIGIBLE","NO")')
    row += 2

    set_label(row, 'STEP 2: CALCULATE ORIGINAL SALE PROCEEDS & GAIN')
    row += 1
    set_label(row, '  Gross proceeds from sale')
    set_formula(row, 100008, '$#,##0')
    row += 1
    set_label(row, '  Cost basis')
    set_formula(row, 100000, '$#,##0')
    row += 1
    set_label(row, '  Realized gain')
    set_formula(row, '=B67-B69', '$#,##0')
    row += 2

    set_label(row, 'STEP 3: APPLY §1045 DEFERRAL')
    row += 1
    set_label(row, '  Amount reinvested in replacement QSBS')
    set_formula(row, 100008, '$#,##0')
    row += 1
    set_label(row, '  Deferred gain')
    set_formula(row, '=IF(B63="YES - §1045 ELIGIBLE",MIN(B71,B76),0)', '$#,##0')
    row += 1
    set_label(row, '  Recognized gain (taxed now)')
    set_formula(row, '=MAX(0,B71-B78)', '$#,##0')
    row += 1
    set_label(row, '  Tax on recognized gain')
    set_formula(row, '=B80*combined_rate', '$#,##0')
    row += 2

    set_label(row, 'STEP 4: CALCULATE ADJUSTED BASIS IN REPLACEMENT')
    row += 1
    set_label(row, '  Cost of replacement stock')
    set_formula(row, '=B76', '$#,##0')
    row += 1
    set_label(row, '  Adjusted basis (cost minus deferred gain)')
    set_formula(row, '=B87-B78', '$#,##0')
    set_result_label(row, '← Lower basis = gain recaptured later')
    row += 1
    set_label(row, '  Shares of replacement acquired')
    set_formula(row, '=B76/Inputs!$B$30', '#,##0')
    row += 2

    set_label(row, 'STEP 5: DETERMINE TACKED HOLDING PERIOD')
    row += 1
    set_label(row, '  Tacked start date (original purchase)')
    set_formula(row, '=serb_date', 'yyyy-mm-dd')
    row += 1
    set_label(row, '  Replacement sale date (example)')
    set_formula(row, '=evecxia_ph2_date', 'yyyy-mm-dd')
    row += 1
    set_label(row, '  Tacked holding period (days)')
    set_formula(row, '=B99-B97', '#,##0')
    row += 1
    set_label(row, '  Tacked holding period (years)')
    set_formula(row, '=B101/365.25', '0.00')
    row += 1
    set_label(row, '  QSBS 5-year met at replacement sale?')
    set_formula(row, '=IF(B101>=1826,"YES","NO")')
    row += 2

    set_label(row, 'STEP 6: CALCULATE REPLACEMENT SALE & §1202 EXCLUSION')
    row += 1
    set_label(row, '  Replacement sale proceeds (example)')
    set_formula(row, 3845020, '$#,##0')
    row += 1
    set_label(row, '  Gain on replacement sale')
    set_formula(row, '=B109-B89', '$#,##0')
    row += 2

    set_label(row, '  §1202 exclusion cap (if eligible)')
    set_formula(row, '=IF(B105="YES",MAX(10000000,10*B69),0)', '$#,##0')
    set_result_label(row, '← Uses ORIGINAL basis')
    row += 1
    set_label(row, '  Excluded gain')
    set_formula(row, '=IF(AND(B105="YES",B59="YES"),MIN(B111,B114),0)', '$#,##0')
    row += 1
    set_label(row, '  Remaining taxable gain')
    set_formula(row, '=MAX(0,B111-B117)', '$#,##0')
    row += 2

    set_label(row, 'STEP 7: CALCULATE FINAL TAX & NET PROCEEDS')
    row += 1
    set_label(row, '  Tax on replacement sale')
    set_formula(row, '=B119*combined_rate', '$#,##0')
    row += 1
    set_label(row, '  Total tax (original + replacement)')
    set_formula(row, '=B82+B123', '$#,##0')
    row += 1
    set_label(row, '  Net after-tax proceeds (full journey)')
    set_formula(row, '=B109-B125', '$#,##0')
    row += 1
    set_label(row, '  MOIC (vs original $100K basis)')
    set_formula(row, '=B127/B69', '0.00x')
    row += 2

    # ============================================================
    # SECTION 4: COMBINED §1045 → §1202 CHAIN LOGIC
    # ============================================================
    set_header(row, 'SECTION 4: COMBINED §1045 → §1202 CHAIN (SUMMARY)')
    row += 1

    set_label(row, 'This section validates the full chain:')
    row += 1
    set_label(row, '1. Original QSBS held > 6 months')
    row += 1
    set_label(row, '2. Sold and reinvested in replacement QSBS within 60 days')
    row += 1
    set_label(row, '3. Gain deferred under §1045')
    row += 1
    set_label(row, '4. Basis in replacement reduced by deferred gain')
    row += 1
    set_label(row, '5. Holding period tacks from original purchase date')
    row += 1
    set_label(row, '6. If tacked period ≥ 5 years at replacement sale → §1202 available')
    row += 1
    set_label(row, '7. §1202 exclusion applies to gain on replacement (up to $10M or 10× ORIGINAL basis)')
    row += 2

    set_label(row, 'Key Rule: §1202 10× basis test uses ORIGINAL QSBS basis, not adjusted replacement basis')
    ws[f'A{row}'].font = Font(name='Calibri', size=11, bold=True, italic=True)
    row += 2

    set_label(row, 'Example Chain Validation:')
    row += 1
    set_label(row, '  Original investment: $100,000 (May 3, 2022)')
    row += 1
    set_label(row, '  Sold at lock-up (March 12, 2026) for ~$100,008')
    row += 1
    set_label(row, '  Gain: $8 → deferred via §1045')
    row += 1
    set_label(row, '  Reinvested $100,008 in Evecxia (April 11, 2026)')
    row += 1
    set_label(row, '  Evecxia adjusted basis: $100,000 ($100,008 cost - $8 deferred)')
    row += 1
    set_label(row, '  Sold Evecxia Oct 15, 2027 for $3,845,020')
    row += 1
    set_label(row, '  Tacked holding: May 3, 2022 → Oct 15, 2027 = 1,991 days (5.45 years) ✓')
    row += 1
    set_label(row, '  Evecxia gain: $3,745,020')
    row += 1
    set_label(row, '  §1202 cap: MAX($10M, 10 × $100K original) = $10M')
    row += 1
    set_label(row, '  Excluded: $3,745,020 (fully excluded, under cap)')
    row += 1
    set_label(row, '  Tax: $0 federal (§1202 exclusion)')
    row += 1
    set_label(row, '  Net: $3,845,020 | MOIC: 38.5x | Effective tax rate: 0%')
    row += 2

    # Freeze panes
    ws.freeze_panes = 'A4'

def main():
    print("Building LBRX QSBS Evecxia Model - Chunk 2: TaxLogic Tab...")

    # Load checkpoint
    wb = load_workbook('/home/user/LB-Pharma-qsbs/LBRX_QSBS_Evecxia_Model_v1_inputs.xlsx')
    print("✓ Loaded checkpoint v1")

    # Build TaxLogic tab
    build_taxlogic_tab(wb)
    print("✓ Built TaxLogic tab (4 sections, ~165 rows)")

    # Save checkpoint
    output_file = '/home/user/LB-Pharma-qsbs/LBRX_QSBS_Evecxia_Model_v2_taxlogic.xlsx'
    wb.save(output_file)
    print(f"✓ Saved checkpoint: {output_file}")
    print("\nChunk 2 complete!")

if __name__ == '__main__':
    main()
