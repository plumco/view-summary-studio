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

with open('all_summaries_detailed.txt', 'w', encoding='utf-8') as out:
    for idx, fname in enumerate(files, 1):
        out.write(f"\n{'='*80}\nFILE {idx}: {fname}\n{'='*80}\n")
        wb_f = openpyxl.load_workbook(fname, data_only=False)
        wb_v = openpyxl.load_workbook(fname, data_only=True)
        
        ws_f = wb_f['Summary Sheet']
        ws_v = wb_v['Summary Sheet']
        
        out.write(f"Sheet dimensions: max_row={ws_f.max_row}, max_column={ws_f.max_column}\n")
        out.write(f"Merged cell ranges:\n")
        for m in sorted(ws_f.merged_cells.ranges, key=lambda x: (x.min_row, x.min_col)):
            out.write(f"  {m}\n")
        
        out.write("\nAll non-empty cells:\n")
        for r in range(1, ws_f.max_row + 1):
            row_items = []
            for c in range(1, ws_f.max_column + 1):
                f_val = ws_f.cell(row=r, column=c).value
                v_val = ws_v.cell(row=r, column=c).value
                if f_val is not None or v_val is not None:
                    col_let = get_column_letter(c)
                    coord = f"{col_let}{r}"
                    if str(f_val).startswith('='):
                        row_items.append(f"{coord}: [FORMULA={f_val} | VAL={v_val}]")
                    else:
                        row_items.append(f"{coord}: {v_val}")
            if row_items:
                out.write(f"Row {r:2d} -> " + " | ".join(row_items) + "\n")
print("Done writing all_summaries_detailed.txt")
