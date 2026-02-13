#!/usr/bin/env python3
"""
LBRX QSBS Evecxia Model - Chunk 3: Timeline + Comps Tabs
Timeline: Event chronology with QSBS flags
Comps: 4 comp sets (LBRX pos/neg, Evecxia pos/neg) with pre-loaded + verified data
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime, timedelta

def build_timeline_tab(wb):
    """Build Timeline tab with date-anchored events and QSBS flags"""

    ws = wb['Timeline']

    # Set column widths
    ws.column_dimensions['A'].width = 15  # Date
    ws.column_dimensions['B'].width = 50  # Event
    ws.column_dimensions['C'].width = 18  # Days from Series B
    ws.column_dimensions['D'].width = 20  # QSBS Status
    ws.column_dimensions['E'].width = 40  # Notes

    # Headers
    headers = ['Date', 'Event', 'Days from Series B', 'QSBS Status', 'Notes']
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.value = header
        cell.font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Events (row by row)
    events = [
        ('=serb_date', 'Series B Closes → QSBS Clock Starts', '=A2-serb_date', 'Day 0', 'QSBS holding period begins'),
        ('=serb_date+180', '6-Month §1045 Eligibility Threshold Met', '=A3-serb_date', '6 months', '§1045 rollover now available if sold'),
        ('=Inputs!$B$39', 'LBRX IPO at $15.00/share', '=A4-serb_date', 'IPO event', 'Series B converts to common, heavy dilution'),
        ('=lockup_date', 'Lock-Up Release (180 days post-IPO)', '=A5-serb_date', '~46 months', 'Series B can sell, but < 5 years → NO §1202'),
        ('=evecxia_purchase', 'Evecxia Purchase (§1045 scenario)', '=A6-serb_date', 'Within 60d of lock-up', 'If §1045 rollover executed'),
        ('=qsbs_5yr_date', '5-YEAR QSBS THRESHOLD MET (Series B)', '=A7-serb_date', '1826 days (5 yrs)', '§1202 exclusion NOW available'),
        ('=Inputs!$B$49', 'Evecxia IPO (estimated)', '=A8-serb_date', 'Pre-readout', 'If Evecxia goes public Q2 2027'),
        ('=ph3_date', 'LBRX Phase 3 Topline (estimated)', '=A9-serb_date', '~5.4 years', 'CRITICAL: After May 3, 2027 → §1202 available'),
        ('=evecxia_ph2_date', 'Evecxia Phase 2 OCD Readout (estimated)', '=A10-serb_date', '~5.45 years', '§1045→§1202 chain value inflection')
    ]

    for row_num, (date_formula, event, days_formula, qsbs_status, notes) in enumerate(events, 2):
        ws[f'A{row_num}'] = date_formula
        ws[f'A{row_num}'].number_format = 'yyyy-mm-dd'

        ws[f'B{row_num}'] = event
        ws[f'B{row_num}'].font = Font(name='Calibri', size=10)

        ws[f'C{row_num}'] = days_formula
        ws[f'C{row_num}'].number_format = '#,##0'
        ws[f'C{row_num}'].alignment = Alignment(horizontal='right')

        ws[f'D{row_num}'] = qsbs_status
        ws[f'D{row_num}'].alignment = Alignment(horizontal='center')

        ws[f'E{row_num}'] = notes
        ws[f'E{row_num}'].font = Font(name='Calibri', size=9, italic=True)
        ws[f'E{row_num}'].alignment = Alignment(wrap_text=True)

    # Add key flag formulas below table
    row = 13
    ws[f'A{row}'] = 'KEY FLAGS (formula-driven):'
    ws[f'A{row}'].font = Font(name='Calibri', size=11, bold=True)
    row += 1

    flags = [
        ('Lock-up sale before 5-year QSBS?', '=IF(lockup_date<qsbs_5yr_date,"YES - NO §1202 AVAILABLE","NO - §1202 AVAILABLE")', 'Critical for tax planning'),
        ('Phase 3 readout after QSBS 5-year?', '=IF(ph3_date>qsbs_5yr_date,"YES - §1202 AVAILABLE","NO - TOO EARLY")', 'Can hold through Ph3 with §1202'),
        ('§1045 holding > 6 months at lock-up?', '=IF((lockup_date-serb_date)>180,"YES - §1045 ELIGIBLE","NO")', 'Series B meets 6-month requirement'),
        ('Evecxia purchase within 60 days?', '=IF((evecxia_purchase-lockup_date)<=60,"YES - WITHIN WINDOW","NO - TOO LATE")', 'Must reinvest within 60 days'),
        ('Tacked holding at Evecxia sale (days)', '=evecxia_ph2_date-serb_date', 'Holding period tacks from Series B date'),
        ('Tacked holding at Evecxia sale (years)', '=(evecxia_ph2_date-serb_date)/365.25', '> 5 years → §1202 on Evecxia sale')
    ]

    for label, formula, note in flags:
        ws[f'A{row}'] = label
        ws[f'A{row}'].font = Font(name='Calibri', size=10)

        ws[f'B{row}'] = formula
        ws[f'B{row}'].alignment = Alignment(horizontal='left')

        ws[f'E{row}'] = note
        ws[f'E{row}'].font = Font(name='Calibri', size=9, italic=True)

        row += 1

    # Set number formats for numeric flag formulas
    ws['B18'].number_format = '#,##0'  # Days
    ws['B19'].number_format = '0.00'    # Years

    # Freeze panes
    ws.freeze_panes = 'A2'


def build_comps_tab(wb):
    """Build Comps tab with 4 sections based on pre-loaded + verified data"""

    ws = wb['Comps']

    # Set column widths
    widths = {'A': 25, 'B': 10, 'C': 25, 'D': 15, 'E': 20, 'F': 15, 'G': 15, 'H': 15, 'I': 12, 'J': 50, 'K': 60}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    def add_section_header(row, title):
        ws[f'A{row}'] = title
        ws[f'A{row}'].font = Font(name='Calibri', size=12, bold=True, color='FFFFFF')
        ws[f'A{row}'].fill = PatternFill(start_color='2E75B6', end_color='2E75B6', fill_type='solid')
        ws.merge_cells(f'A{row}:K{row}')
        return row + 1

    def add_comp_table_headers(row):
        headers = ['Company', 'Ticker', 'Indication', 'Stage', 'Event', 'Event Date', 'Mkt Cap ($M)', 'EV ($M)', '% Move', 'Rationale', 'Source']
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col_num)
            cell.value = header
            cell.font = Font(name='Calibri', size=10, bold=True)
            cell.fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        return row + 1

    def add_comp_row(row, data):
        for col_num, value in enumerate(data, 1):
            cell = ws.cell(row=row, column=col_num)
            cell.value = value
            cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
            if col_num in [7, 8]:  # Mkt Cap, EV
                cell.number_format = '#,##0'
            elif col_num == 9:  # % Move
                if isinstance(value, str) and value != 'N/A':
                    cell.value = value
                elif isinstance(value, (int, float)):
                    cell.value = value / 100
                    cell.number_format = '0%'
        return row + 1

    row = 1

    # ============================================================
    # COMP SET 1: LBRX POST-PHASE 3 POSITIVE
    # ============================================================
    row = add_section_header(row, 'COMP SET 1: LBRX POST-PHASE 3 POSITIVE READOUT VALUATION')
    row = add_comp_table_headers(row)

    comps_lbrx_pos = [
        ('Karuna Therapeutics', 'KRTX', 'Schizophrenia (KarXT)', 'FDA review stage', 'BMY Acquisition Announced', 'Dec 22 2023', 14000, 12700, 53, 'Premier schizophrenia comp, 330 per share', 'BMY press release Dec 22, 2023'),
        ('Intra-Cellular Therapies', 'ITCI', 'Schizophrenia + Bipolar (Caplyta)', 'FDA approved, commercial', 'JNJ Acquisition Announced', 'Jan 13 2025', 14600, 14600, 39, 'CNS antipsychotic, 132 per share, 2024 revenue 481M', 'JNJ press release Jan 13, 2025'),
        ('Cerevel Therapeutics', 'CERE', 'Schizophrenia (emraclidine)', 'Phase 2 (pre-data at acquisition)', 'AbbVie Acquisition Announced', 'Dec 6 2023', 8700, 8700, 22, 'CAUTION: Emraclidine Ph2 failed Nov 2024, 3.5B impairment', 'AbbVie press release Dec 6, 2023'),
        ('Neurocrine Biosciences', 'NBIX', 'Tardive dyskinesia (Ingrezza)', 'Commercial (approved 2017)', 'Current Market Cap', 'Feb 13 2026', 12300, 12300, 'N/A', 'CNS commercial comp, long-term value ceiling', 'Yahoo Finance Feb 2026 (updated from 15B)'),
        ('Acadia Pharmaceuticals', 'ACAD', "Parkinson's psychosis (Nuplazid)", 'Commercial (approved 2016)', 'Current Market Cap', 'Feb 13 2026', 4300, 4300, 'N/A', 'Psychiatric niche indication, single-product CNS', 'Yahoo Finance Feb 2026 (updated from 3.8B)')
    ]

    for comp_data in comps_lbrx_pos:
        row = add_comp_row(row, comp_data)

    row += 1
    ws[f'A{row}'] = 'DERIVED LBRX POSITIVE READOUT RANGE:'
    ws[f'A{row}'].font = Font(name='Calibri', size=11, bold=True)
    row += 1

    ws[f'A{row}'] = 'Low'
    ws[f'B{row}'] = '$1.5B mkt cap (~$50/share)'
    ws[f'J{row}'] = 'Discount to Cerevel emraclidine-only; single-asset Phase 3 positive but pre-NDA'
    row += 1
    ws[f'A{row}'] = 'Base'
    ws[f'B{row}'] = '$3.0B mkt cap (~$100/share)'
    ws[f'J{row}'] = 'Between ITCI at approval ($1.8B historical) and Cerevel ($8.7B); accounts for positive Phase 3 + pipeline'
    row += 1
    ws[f'A{row}'] = 'High'
    ws[f'B{row}'] = '$5.0B mkt cap (~$165/share)'
    ws[f'J{row}'] = 'Approaches Karuna pre-deal valuation; assumes very strong data + M&A premium expectations'

    row += 2

    # ============================================================
    # COMP SET 2: LBRX POST-PHASE 3 NEGATIVE
    # ============================================================
    row = add_section_header(row, 'COMP SET 2: LBRX POST-PHASE 3 NEGATIVE READOUT / FAILURE')
    row = add_comp_table_headers(row)

    comps_lbrx_neg = [
        ('Minerva Neurosciences', 'NERV', 'Schizophrenia negative symptoms', 'Phase 3 failure', 'Roluperidone Ph3 Failure', 'May 29 2020', 100, 100, -81, 'Direct schizophrenia Ph3 failure, 550M to 100M', 'Seeking Alpha, Minerva PR May 29, 2020'),
        ('Sage Therapeutics', 'SAGE', 'Major depressive disorder', 'Phase 3 failure', 'SAGE-217 MOUNTAIN Failure', 'Dec 5 2019', 3700, 3700, -60, 'CNS psych Ph3 failure in depression, 8B to 3.7B', 'Sage press release Dec 5, 2019'),
    ]

    for comp_data in comps_lbrx_neg:
        row = add_comp_row(row, comp_data)

    row += 1
    ws[f'A{row}'] = 'DERIVED LBRX NEGATIVE READOUT RANGE:'
    ws[f'A{row}'].font = Font(name='Calibri', size=11, bold=True)
    row += 1

    ws[f'A{row}'] = 'Low'
    ws[f'B{row}'] = '$100M mkt cap (~$3.30/share)'
    ws[f'J{row}'] = 'Minerva-style wipeout; total loss of schizophrenia thesis, bipolar program too early'
    row += 1
    ws[f'A{row}'] = 'Base'
    ws[f'B{row}'] = '$200M mkt cap (~$6.60/share)'
    ws[f'J{row}'] = 'Cash floor (~$300M+ cash) minus burn; bipolar program provides option value'
    row += 1
    ws[f'A{row}'] = 'High'
    ws[f'B{row}'] = '$350M mkt cap (~$11.50/share)'
    ws[f'J{row}'] = 'Modest decline; if failure is dose-related or borderline, market may give credit for re-try + bipolar pipeline'

    row += 2

    # ============================================================
    # COMP SET 3: EVECXIA PRE-PHASE 2 READOUT / IPO
    # ============================================================
    row = add_section_header(row, 'COMP SET 3: EVECXIA PRE-PHASE 2 READOUT / IPO VALUATION')
    row = add_comp_table_headers(row)

    comps_evecxia_ipo = [
        ('Karuna Therapeutics', 'KRTX', 'Schizophrenia (KarXT)', 'Pre-Phase 2 data', 'IPO', 'Jul 1 2019', 414, 414, 'N/A', 'Novel mechanism psych company at IPO, Phase 2 data came Nov 2019', 'Karuna S-1, Streetwise Reports Nov 2019'),
        ('Axsome Therapeutics', 'AXSM', 'Multi-indication CNS platform', 'Early clinical', 'IPO', 'Jan 1 2015', 155, 155, 'N/A', 'Psych multi-indication platform, IPOd small, re-rated to 9.4B by 2026', 'StockAnalysis.com'),
    ]

    for comp_data in comps_evecxia_ipo:
        row = add_comp_row(row, comp_data)

    row += 1
    ws[f'A{row}'] = 'DERIVED EVECXIA IPO RANGE (Q2 2027):'
    ws[f'A{row}'].font = Font(name='Calibri', size=11, bold=True)
    row += 1

    ws[f'A{row}'] = 'Low'
    ws[f'B{row}'] = '$150M post-money'
    ws[f'J{row}'] = 'Conservative; Phase 1 complete, OCD Phase 2 enrolling, no data yet. Reflects Series A step-up.'
    row += 1
    ws[f'A{row}'] = 'Base'
    ws[f'B{row}'] = '$400M post-money'
    ws[f'J{row}'] = 'In line with Karuna at IPO; assumes strong investor interest in serotonin platform + OCD thesis'
    row += 1
    ws[f'A{row}'] = 'High'
    ws[f'B{row}'] = '$750M post-money'
    ws[f'J{row}'] = 'Assumes pre-data hype cycle comparable to Karuna pre-Phase 2; M&A interest creates premium'

    row += 2

    # ============================================================
    # COMP SET 4: EVECXIA POST-PHASE 2 (POSITIVE VS NEGATIVE)
    # ============================================================
    row = add_section_header(row, 'COMP SET 4: EVECXIA POST-PHASE 2 POSITIVE VS NEGATIVE OUTCOMES')
    row = add_comp_table_headers(row)

    comps_evecxia_readout = [
        ('Karuna Therapeutics', 'KRTX', 'Schizophrenia (KarXT)', 'Phase 2 positive', 'Phase 2 Data Announced', 'Nov 18 2019', 2000, 2000, 375, 'Premier Phase 2 psych readout comp, 11.6-pt PANSS reduction, $414M → $2B in 1 day', 'Karuna press release Nov 18, 2019'),
        ('Sage Therapeutics', 'SAGE', 'Postpartum depression', 'Phase 2 success', 'Phase 2 PPD Success', 'Jan 1 2018', 1500, 1500, 'N/A', 'Phase 2 psych success trajectory, rose from ~$80 to ~$190 in 2019', 'Sage press releases 2018-2019'),
        ('Axsome Therapeutics', 'AXSM', 'Major depressive disorder', 'Phase 3 MDD success', 'AXS-05 Ph3 MDD Success', 'Dec 1 2019', 3800, 3800, 'N/A', 'Depression Phase 3 success, grew from ~$155M at IPO to ~$3.8B end 2019', 'Axsome press releases'),
    ]

    for comp_data in comps_evecxia_readout:
        row = add_comp_row(row, comp_data)

    row += 1
    ws[f'A{row}'] = 'DERIVED EVECXIA POST-PHASE 2 POSITIVE RANGE:'
    ws[f'A{row}'].font = Font(name='Calibri', size=11, bold=True)
    row += 1

    ws[f'A{row}'] = 'Low'
    ws[f'B{row}'] = '$800M'
    ws[f'J{row}'] = 'Modest positive; OCD data supportive but not blockbuster'
    row += 1
    ws[f'A{row}'] = 'Base'
    ws[f'B{row}'] = '$1.5B'
    ws[f'J{row}'] = 'Clearly positive Phase 2; validates serotonin platform; M&A interest intensifies'
    row += 1
    ws[f'A{row}'] = 'High'
    ws[f'B{row}'] = '$3.0B'
    ws[f'J{row}'] = 'Karuna-style blowout; data exceeds expectations; bidding war / strategic premium'
    row += 2

    ws[f'A{row}'] = 'DERIVED EVECXIA POST-PHASE 2 NEGATIVE RANGE:'
    ws[f'A{row}'].font = Font(name='Calibri', size=11, bold=True)
    row += 1

    ws[f'A{row}'] = 'Low'
    ws[f'B{row}'] = '$50M'
    ws[f'J{row}'] = 'Cash-floor wipeout; OCD thesis invalidated'
    row += 1
    ws[f'A{row}'] = 'Base'
    ws[f'B{row}'] = '$120M'
    ws[f'J{row}'] = 'Failed primary but secondary signals; platform retains option value in other indications'
    row += 1
    ws[f'A{row}'] = 'High'
    ws[f'B{row}'] = '$250M'
    ws[f'J{row}'] = 'Borderline miss; dose/endpoint issue; still fundable'

    # Freeze panes
    ws.freeze_panes = 'A3'


def main():
    print("Building LBRX QSBS Evecxia Model - Chunk 3: Timeline + Comps Tabs...")

    # Load checkpoint
    wb = load_workbook('/home/user/LB-Pharma-qsbs/LBRX_QSBS_Evecxia_Model_v2_taxlogic.xlsx')
    print("✓ Loaded checkpoint v2")

    # Build Timeline tab
    build_timeline_tab(wb)
    print("✓ Built Timeline tab (9 events + 6 key flags)")

    # Build Comps tab
    build_comps_tab(wb)
    print("✓ Built Comps tab (4 comp sets: LBRX pos/neg, Evecxia IPO/readout)")

    # Save checkpoint
    output_file = '/home/user/LB-Pharma-qsbs/LBRX_QSBS_Evecxia_Model_v3_timeline_comps.xlsx'
    wb.save(output_file)
    print(f"✓ Saved checkpoint: {output_file}")
    print("\nChunk 3 complete!")

if __name__ == '__main__':
    main()
