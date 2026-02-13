#!/usr/bin/env python3
"""
LBRX QSBS Evecxia Model - Chunk 1: Workbook Setup + Inputs Tab
Creates the base workbook with 8 tabs and builds the Inputs tab per exact cell map.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, Fill, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from datetime import date, datetime, timedelta

def create_workbook_structure():
    """Create workbook with 8 tabs"""
    wb = Workbook()

    # Remove default sheet
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    # Create all 8 tabs
    tab_names = ['README', 'Inputs', 'Timeline', 'Comps', 'Scenarios', 'TaxLogic', 'Outputs', 'Sources']
    for name in tab_names:
        wb.create_sheet(title=name)

    return wb

def setup_styles():
    """Define reusable styles"""
    styles = {
        'header_dark': {
            'font': Font(name='Calibri', size=11, bold=True, color='FFFFFF'),
            'fill': PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid'),
            'alignment': Alignment(horizontal='left', vertical='center')
        },
        'header_medium': {
            'font': Font(name='Calibri', size=11, bold=True, color='FFFFFF'),
            'fill': PatternFill(start_color='2E75B6', end_color='2E75B6', fill_type='solid'),
            'alignment': Alignment(horizontal='left', vertical='center')
        },
        'input_cell': {
            'fill': PatternFill(start_color='DCE6F1', end_color='DCE6F1', fill_type='solid'),
            'alignment': Alignment(horizontal='right', vertical='center')
        },
        'formula_cell': {
            'fill': PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid'),
            'alignment': Alignment(horizontal='right', vertical='center')
        },
        'label_cell': {
            'font': Font(name='Calibri', size=11),
            'alignment': Alignment(horizontal='left', vertical='center')
        },
        'section_header': {
            'font': Font(name='Calibri', size=11, bold=True),
            'fill': PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid'),
            'alignment': Alignment(horizontal='left', vertical='center')
        }
    }
    return styles

def build_inputs_tab(ws, styles):
    """Build Inputs tab per exact cell map from prompt"""

    # Set column widths
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 20

    # Row 1-2: Headers
    ws['A1'] = 'LBRX QSBS & Evecxia Scenario Model'
    ws['A1'].font = Font(name='Calibri', size=14, bold=True)

    ws['A2'] = 'Inputs & Assumptions'
    ws['A2'].font = Font(name='Calibri', size=12, bold=True)

    # Helper function to set cell
    def set_row(row_num, label, value=None, is_input=True, is_formula=False, number_format=None):
        ws[f'A{row_num}'] = label
        ws[f'A{row_num}'].font = styles['label_cell']['font']
        ws[f'A{row_num}'].alignment = styles['label_cell']['alignment']

        if value is not None or is_formula:
            cell = ws[f'B{row_num}']
            if is_formula:
                cell.value = value  # Value is the formula string
            else:
                cell.value = value

            if is_input:
                cell.fill = styles['input_cell']['fill']
            else:
                cell.fill = styles['formula_cell']['fill']

            cell.alignment = Alignment(horizontal='right', vertical='center')

            if number_format:
                cell.number_format = number_format

    # Section headers helper
    def set_section(row_num, label):
        ws[f'A{row_num}'] = label
        ws[f'A{row_num}'].font = Font(name='Calibri', size=11, bold=True)
        ws[f'A{row_num}'].fill = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')

    # Row 3: Blank

    # Row 4-12: INVESTOR TAX PROFILE
    set_section(4, 'INVESTOR TAX PROFILE')
    set_row(5, 'Federal LTCG Rate', 0.20, True, False, '0.0%')
    set_row(6, 'NIIT (3.8%) Applies?', 'YES', True, False)
    set_row(7, 'NIIT Rate', '=IF(B6="YES",0.038,0)', False, True, '0.0%')
    set_row(8, 'Effective Federal Rate', '=B5+B7', False, True, '0.0%')
    set_row(9, 'State Tax Rate', 0.00, True, False, '0.0%')
    set_row(10, 'Combined Tax Rate', '=B8+B9', False, True, '0.0%')
    set_row(11, 'Federal Ordinary Income Rate', 0.37, True, False, '0.0%')

    # Row 13-19: SERIES B INVESTMENT
    set_section(13, 'SERIES B INVESTMENT')
    set_row(14, 'Series B Invest Date', datetime(2022, 5, 3), True, False, 'yyyy-mm-dd')
    set_row(15, 'Series B Cost Basis ($)', 100000, True, False, '$#,##0')
    set_row(16, 'Series B Cost Per Share', 24.00, True, False, '$#,##0.00')
    set_row(17, 'Series B Shares Acquired', '=B15/B16', False, True, '#,##0')
    set_row(18, 'Series B QSBS Eligible', 'YES', False, False)
    set_row(19, 'QSBS 5-Year Completion Date', '=B14+1826', False, True, 'yyyy-mm-dd')  # 5 years = 1826 days (5*365.25 rounded)

    # Row 21-26: SERIES C INVESTMENT
    set_section(21, 'SERIES C INVESTMENT (optional)')
    set_row(22, 'Series C Invest Date', None, True, False, 'yyyy-mm-dd')
    set_row(23, 'Series C Cost Basis ($)', None, True, False, '$#,##0')
    set_row(24, 'Series C Cost Per Share', 36.44, True, False, '$#,##0.00')
    set_row(25, 'Series C Shares Acquired', '=IF(B23="","",B23/B24)', False, True, '#,##0')
    set_row(26, 'Series C QSBS Eligible', 'NO', False, False)

    # Row 28-36: EVECXIA INVESTMENT
    set_section(28, 'EVECXIA INVESTMENT (§1045)')
    set_row(29, 'Evecxia Purchase Date', datetime(2026, 4, 11), True, False, 'yyyy-mm-dd')
    set_row(30, 'Evecxia Share Price', 0.1069, True, False, '$#,##0.0000')
    set_row(31, 'Amount Reinvested', '=B15', False, True, '$#,##0')  # Linked from Series B basis for now
    set_row(32, 'Evecxia Shares Acquired', '=B31/B30', False, True, '#,##0')
    set_row(33, 'Evecxia QSBS Eligible', 'YES', True, False)
    set_row(34, 'Evecxia Pre-Money Valuation', 14000000, True, False, '$#,##0')
    set_row(35, 'Evecxia Post-Money Valuation', 30000000, True, False, '$#,##0')
    set_row(36, 'Evecxia Total Shares at SerA Close', '=B35/B30', False, True, '#,##0')

    # Row 38-44: MARKET / IPO DETAILS
    set_section(38, 'MARKET / IPO DETAILS')
    set_row(39, 'IPO Date', datetime(2025, 9, 12), True, False, 'yyyy-mm-dd')
    set_row(40, 'IPO Price', 15.00, True, False, '$#,##0.00')
    set_row(41, 'Lock-Up Release Date', datetime(2026, 3, 12), True, False, 'yyyy-mm-dd')  # 180 days from IPO
    set_row(42, 'Current / Lock-Up Price', 24.22, True, False, '$#,##0.00')  # From research: $24.22
    set_row(43, 'Shares Outstanding (post-IPO)', 25300000, True, False, '#,##0')
    set_row(44, 'Shares Outstanding (post-PIPE)', 30000000, True, False, '#,##0')

    # Row 46-48: CLINICAL TIMELINE
    set_section(46, 'CLINICAL TIMELINE')
    set_row(47, 'Phase 3 Topline Expected', datetime(2027, 10, 1), True, False, 'yyyy-mm-dd')
    set_row(48, 'Evecxia Phase 2 Readout', datetime(2027, 10, 15), True, False, 'yyyy-mm-dd')
    set_row(49, 'Evecxia IPO Target', datetime(2027, 6, 1), True, False, 'yyyy-mm-dd')

    # Row 51-65: SCENARIO VALUATIONS
    set_section(51, 'SCENARIO VALUATIONS (from Comps)')
    set_row(52, 'LBRX Price — Ph3 Positive (Low)', 50.00, True, False, '$#,##0.00')
    set_row(53, 'LBRX Price — Ph3 Positive (Base)', 100.00, True, False, '$#,##0.00')
    set_row(54, 'LBRX Price — Ph3 Positive (High)', 165.00, True, False, '$#,##0.00')
    set_row(55, 'LBRX Price — Ph3 Negative (Low)', 3.30, True, False, '$#,##0.00')
    set_row(56, 'LBRX Price — Ph3 Negative (Base)', 6.60, True, False, '$#,##0.00')
    set_row(57, 'LBRX Price — Ph3 Negative (High)', 11.50, True, False, '$#,##0.00')
    set_row(58, 'Evecxia Val — Ph2 Positive (Low)', 800000000, True, False, '$#,##0')
    set_row(59, 'Evecxia Val — Ph2 Positive (Base)', 1500000000, True, False, '$#,##0')
    set_row(60, 'Evecxia Val — Ph2 Positive (High)', 3000000000, True, False, '$#,##0')
    set_row(61, 'Evecxia Val — Ph2 Negative (Low)', 50000000, True, False, '$#,##0')
    set_row(62, 'Evecxia Val — Ph2 Negative (Base)', 120000000, True, False, '$#,##0')
    set_row(63, 'Evecxia Val — Ph2 Negative (High)', 250000000, True, False, '$#,##0')
    set_row(64, 'Evecxia IPO Valuation', 400000000, True, False, '$#,##0')
    set_row(65, 'Evecxia Shares at Readout (est.)', 365000000, True, False, '#,##0')

    # Freeze panes at row 3
    ws.freeze_panes = 'A3'

def define_named_ranges(wb):
    """Define named ranges for key inputs"""
    from openpyxl.workbook.defined_name import DefinedName

    # Define named ranges per cell map
    named_ranges = {
        'fed_ltcg_rate': 'Inputs!$B$5',
        'niit_rate': 'Inputs!$B$7',
        'eff_fed_rate': 'Inputs!$B$8',
        'state_rate': 'Inputs!$B$9',
        'combined_rate': 'Inputs!$B$10',
        'serb_date': 'Inputs!$B$14',
        'serb_basis': 'Inputs!$B$15',
        'serb_shares': 'Inputs!$B$17',
        'qsbs_5yr_date': 'Inputs!$B$19',
        'lockup_date': 'Inputs!$B$41',
        'lockup_price': 'Inputs!$B$42',
        'ph3_date': 'Inputs!$B$47',
        'evecxia_ph2_date': 'Inputs!$B$48',
        'evecxia_price': 'Inputs!$B$30',
        'evecxia_purchase': 'Inputs!$B$29'
    }

    for name, ref in named_ranges.items():
        defn = DefinedName(name, attr_text=ref)
        wb.defined_names[name] = defn

def main():
    print("Building LBRX QSBS Evecxia Model - Chunk 1: Inputs Tab...")

    # Create workbook
    wb = create_workbook_structure()
    print("✓ Created workbook with 8 tabs")

    # Setup styles
    styles = setup_styles()
    print("✓ Defined styles")

    # Build Inputs tab
    ws_inputs = wb['Inputs']
    build_inputs_tab(ws_inputs, styles)
    print("✓ Built Inputs tab (65 rows)")

    # Define named ranges
    define_named_ranges(wb)
    print("✓ Defined 16 named ranges")

    # Save checkpoint
    output_file = '/home/user/LB-Pharma-qsbs/LBRX_QSBS_Evecxia_Model_v1_inputs.xlsx'
    wb.save(output_file)
    print(f"✓ Saved checkpoint: {output_file}")
    print("\nChunk 1 complete!")

if __name__ == '__main__':
    main()
