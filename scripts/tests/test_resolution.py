import os
import openpyxl
from openpyxl.utils import get_column_letter

ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
RAW_BOQ_DIR = os.path.join(ROOT_DIR, "data", "raw_boq")

files = [
    ('File 1: Avinea', 'Huliot BOQ Summary AVINEA BY SIGMA 13.08.2026 - 09.10.2026.xlsx'),
    ('File 2: BDR Hospital', 'Huliot BOQ Summary BDR Hospital_08.07.2026.xlsx'),
    ('File 3: Empire', 'Huliot BOQ Summary For Empire Developers 22.10.2026 - 31.10.2026.xlsx'),
    ('File 4: Supernova', 'Huliot BOQ Summary For Supernova by Kanifnath Project, 13.07.2026.xlsx'),
    ('File 5: Tribeca', 'Huliot BOQ Summary Tribeca Lalbaug_10.12.2025- 02.11.2026 (1).xlsx'),
    ('File 6: CASA Terazza', 'Huliot BOQ Summary_CASA Terazza_04.11.2026.xlsx')
]

def find_total_cell(ws, option_name):
    # 1. Find the total row (look for 'No of Floor' in col A, or a row with 'Total' / SUM formula)
    total_row = None
    for r in range(15, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v and 'no of floor' in str(v).lower():
            total_row = r
            break
            
    if not total_row:
        # fallback: find row where column A is empty but has formulas with SUM in other columns
        for r in range(15, ws.max_row + 1):
            for c in range(1, min(ws.max_column+1, 50)):
                f = ws.cell(row=r, column=c).value
                if str(f).startswith('=SUM('):
                    total_row = r
                    break
            if total_row:
                break
                
    # 2. Find the column for this option
    # Look in rows 11, 12, 13 for option_name, then find the 'Total' column at row 14
    opt_col = None
    # We may need to check merged cells
    for c in range(1, ws.max_column + 1):
        # Check if option_name matches header
        hdr_val = None
        for hr in [11, 12, 13]:
            v = ws.cell(row=hr, column=c).value
            if not v:
                for m in ws.merged_cells.ranges:
                    if m.min_row <= hr <= m.max_row and m.min_col <= c <= m.max_col:
                        v = ws.cell(row=m.min_row, column=m.min_col).value
                        break
            if v and option_name.lower() in str(v).lower():
                hdr_val = str(v)
                break
        
        if hdr_val:
            # Under this option block, find the column that has 'Total' in row 14
            sub_val = ws.cell(row=14, column=c).value
            if sub_val and str(sub_val).strip().lower() == 'total':
                opt_col = get_column_letter(c)
                break
                
    return total_row, opt_col

print("TESTING DYNAMIC RESOLUTION OF SOURCE CELLS:")
for label, fname in files:
    print(f"\n--- {label} ---")
    wb = openpyxl.load_workbook(os.path.join(RAW_BOQ_DIR, fname), data_only=False)
    for sname in wb.sheetnames:
        if 'Summary' in sname or 'Design' in sname or sname == 'Typical Floor BOQ ':
            continue
        ws = wb[sname]
        for opt in ['HT PRO', 'Ultra Silent', 'Hybrid', 'Internal - Ultra Silent', 'Drain Pro', 'Decilo']:
            r, c = find_total_cell(ws, opt)
            if c:
                cell_ref = f"'{sname}'!${c}${r}"
                val = ws[f"{c}{r}"].value
                print(f"  Option '{opt}' in '{sname}' -> {cell_ref} [Formula={val}]")
