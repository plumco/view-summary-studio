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

# Let's inspect the target cells of the formulas in each file
# To see what is at those rows and columns in the referenced sheets!

def inspect_file_references(fname):
    print(f"\n=======================================================")
    print(f"FILE: {fname}")
    print(f"=======================================================")
    wb_f = openpyxl.load_workbook(fname, data_only=False)
    wb_v = openpyxl.load_workbook(fname, data_only=True)
    
    ws_summary = wb_f['Summary Sheet']
    
    # Collect all formula references to other sheets
    import re
    cross_refs = set()
    for r in range(1, ws_summary.max_row + 1):
        for c in range(1, ws_summary.max_column + 1):
            val = ws_summary.cell(row=r, column=c).value
            if str(val).startswith('='):
                # match Sheet!Cell
                matches = re.findall(r"'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+)", str(val))
                for s, col, row in matches:
                    cross_refs.add((s.strip("'"), col, int(row), get_column_letter(c) + str(r)))
                # also match range
                range_matches = re.findall(r"'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", str(val))
                for s, c1, r1, c2, r2 in range_matches:
                    cross_refs.add((s.strip("'"), f"{c1}:{c2}", f"{r1}:{r2}", get_column_letter(c) + str(r)))
                    
    print("Cross references from Summary Sheet:")
    for ref in sorted(cross_refs, key=lambda x: (x[0], str(x[1]))):
        sname, col, row, src_cell = ref
        print(f"  Summary {src_cell} -> '{sname}' Col={col}, Row={row}")
        if sname in wb_v.sheetnames:
            ws_target = wb_v[sname]
            if isinstance(row, int):
                # Check target cell value and header
                cell_val = ws_target[f"{col}{row}"].value
                # Let's find header for this column (rows 1-15)
                headers = []
                for hr in range(1, min(15, row)):
                    hv = ws_target.cell(row=hr, column=openpyxl.utils.column_index_from_string(col)).value
                    if hv is not None:
                        headers.append(f"R{hr}:{hv}")
                # check row label (cols A-E)
                row_labels = []
                for hc in range(1, 6):
                    rv = ws_target.cell(row=row, column=hc).value
                    if rv is not None:
                        row_labels.append(f"C{get_column_letter(hc)}:{rv}")
                print(f"    Target Val = {cell_val} | Row Labels = {row_labels} | Col Headers = {headers}")

for f in files:
    inspect_file_references(f)
