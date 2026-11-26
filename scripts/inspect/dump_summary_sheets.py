import openpyxl
from openpyxl.utils import get_column_letter

files = [
    'Huliot BOQ Summary AVINEA BY SIGMA 13.08.2026 - 09.10.2026.xlsx',
    'Huliot BOQ Summary BDR Hospital_08.07.2026.xlsx',
    'Huliot BOQ Summary For Empire Developers 22.10.2026 - 31.10.2026.xlsx',
    'Huliot BOQ Summary For Supernova by Kanifnath Project, 13.07.2026.xlsx',
    'Huliot BOQ Summary Tribeca Lalbaug_10.12.2025- 02.11.2026 (1).xlsx',
    'Huliot BOQ Summary_CASA Terazza_04.11.2026.xlsx'
]

for idx, fname in enumerate(files, 1):
    print(f"\n{'='*70}\nFILE {idx}: {fname}\n{'='*70}")
    wb_formula = openpyxl.load_workbook(fname, data_only=False)
    wb_val = openpyxl.load_workbook(fname, data_only=True)
    
    ws_f = wb_formula['Summary Sheet']
    ws_v = wb_val['Summary Sheet']
    
    print(f"Summary Sheet max_row={ws_f.max_row}, max_column={ws_f.max_column}")
    print(f"Merged ranges: {[str(m) for m in ws_f.merged_cells.ranges]}")
    
    # Let's inspect row by row
    for r in range(1, ws_f.max_row + 1):
        row_content = []
        for c in range(1, ws_f.max_column + 1):
            val = ws_v.cell(row=r, column=c).value
            formula = ws_f.cell(row=r, column=c).value
            
            if val is not None or formula is not None:
                col_let = get_column_letter(c)
                coord = f"{col_let}{r}"
                if str(formula).startswith('='):
                    row_content.append(f"{coord}: [FORMULA: {formula} | VAL: {val}]")
                else:
                    row_content.append(f"{coord}: {val}")
        if row_content:
            print(f"Row {r:2d}: " + " ; ".join(row_content))
