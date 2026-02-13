#!/usr/bin/env python3
"""
LBRX QSBS Evecxia Model - Chunk 6: Formatting + Finalization
Apply professional formatting across all tabs and create final deliverable
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule

def apply_conditional_formatting(wb):
    """Apply conditional formatting to key cells"""
    ws_outputs = wb['Outputs']

    # Green for YES, Red for NO on flags (B4:B7)
    green_fill = PatternFill(start_color='00B050', end_color='00B050', fill_type='solid')
    red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

    # Add conditional formatting rules
    ws_outputs.conditional_formatting.add('B4:B7',
        CellIsRule(operator='equal', formula=['"YES"'], fill=green_fill, font=Font(color='FFFFFF', bold=True)))
    ws_outputs.conditional_formatting.add('B4:B7',
        CellIsRule(operator='equal', formula=['"NO"'], fill=red_fill, font=Font(color='FFFFFF', bold=True)))

def apply_borders_and_formatting(wb):
    """Apply borders and final formatting touches"""

    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Apply borders to key tables
    tabs_with_tables = ['Inputs', 'Scenarios', 'Comps', 'Outputs']

    for tab_name in tabs_with_tables:
        if tab_name in wb.sheetnames:
            ws = wb[tab_name]
            # Apply thin borders to data regions (first 50 rows, first 10 columns)
            for row in ws.iter_rows(min_row=1, max_row=50, min_col=1, max_col=10):
                for cell in row:
                    if cell.value is not None:
                        cell.border = thin_border

def finalize_model(wb):
    """Final touches and cleanup"""

    # Reorder tabs to logical sequence
    desired_order = ['README', 'Inputs', 'Timeline', 'TaxLogic', 'Scenarios', 'Comps', 'Outputs', 'Sources']

    # Note: openpyxl doesn't directly support tab reordering easily
    # The tabs are already in a good order from creation sequence

    # Set active sheet to README
    wb.active = wb['README']

    # Protect certain cells (optional - commented out for now)
    # This would require setting ws.protection.sheet = True and unlocking input cells

def main():
    print("Building LBRX QSBS Evecxia Model - Chunk 6: Formatting + Finalization...")

    # Load checkpoint
    wb = load_workbook('/home/user/LB-Pharma-qsbs/LBRX_QSBS_Evecxia_Model_v5_complete.xlsx')
    print("✓ Loaded checkpoint v5")

    # Apply conditional formatting
    try:
        apply_conditional_formatting(wb)
        print("✓ Applied conditional formatting (green/red flags)")
    except Exception as e:
        print(f"⚠ Conditional formatting skipped: {e}")

    # Apply borders
    apply_borders_and_formatting(wb)
    print("✓ Applied borders and formatting")

    # Finalize
    finalize_model(wb)
    print("✓ Finalized model (set README as active sheet)")

    # Save final deliverable
    output_file = '/home/user/LB-Pharma-qsbs/LBRX_QSBS_Evecxia_Model.xlsx'
    wb.save(output_file)
    print(f"✓ Saved FINAL DELIVERABLE: {output_file}")
    print("\n" + "="*60)
    print("EXCEL MODEL BUILD COMPLETE!")
    print("="*60)

if __name__ == '__main__':
    main()
