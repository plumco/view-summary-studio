import openpyxl
from openpyxl.utils import get_column_letter

files = [
    ('File 1: Avinea', 'Huliot BOQ Summary AVINEA BY SIGMA 13.08.2026 - 09.10.2026.xlsx'),
    ('File 2: BDR Hospital', 'Huliot BOQ Summary BDR Hospital_08.07.2026.xlsx'),
    ('File 3: Empire', 'Huliot BOQ Summary For Empire Developers 22.10.2026 - 31.10.2026.xlsx'),
    ('File 4: Supernova', 'Huliot BOQ Summary For Supernova by Kanifnath Project, 13.07.2026.xlsx'),
    ('File 5: Tribeca', 'Huliot BOQ Summary Tribeca Lalbaug_10.12.2025- 02.11.2026 (1).xlsx'),
    ('File 6: CASA Terazza', 'Huliot BOQ Summary_CASA Terazza_04.11.2026.xlsx')
]

for label, fname in files:
    print(f"\n=================================================================")
    print(f"{label} ({fname})")
    print(f"=================================================================")
    wb_f = openpyxl.load_workbook(fname, data_only=False)
    wb_v = openpyxl.load_workbook(fname, data_only=True)
    ws_f = wb_f['Summary Sheet']
    ws_v = wb_v['Summary Sheet']
    
    # 1. Title Block
    title = ws_v['C2'].value if ws_v['C2'].value else ws_v['B2'].value
    print(f"Title Block: Cell={'C2' if ws_v['C2'].value else 'B2'}, Value='{title}'")
    
    # 2. Main Summary Table
    # Find table header rows (usually rows 3 and 4)
    # Check options in col B or C
    opt_col = 'C' if ws_v['C3'].value == 'Option' else 'B'
    print(f"Main Table Option Column: {opt_col}")
    
    # Headers
    h_r3 = [f"{get_column_letter(c)}3: '{ws_v.cell(row=3, column=c).value}'" for c in range(1, 15) if ws_v.cell(row=3, column=c).value]
    h_r4 = [f"{get_column_letter(c)}4: '{ws_v.cell(row=4, column=c).value}'" for c in range(1, 15) if ws_v.cell(row=4, column=c).value]
    print(f"Header Row 3: {h_r3}")
    print(f"Header Row 4: {h_r4}")
    
    # Options rows (usually rows 4-9 or 5-9)
    print("Options Rows:")
    for r in range(4, 11):
        opt_name = ws_v[f"{opt_col}{r}"].value
        if opt_name and opt_name != 'Discount Structure' and opt_name != 'Ratio':
            row_vals = []
            for c in range(openpyxl.utils.column_index_from_string(opt_col), 15):
                f_val = ws_f.cell(row=r, column=c).value
                v_val = ws_v.cell(row=r, column=c).value
                if f_val is not None:
                    col_let = get_column_letter(c)
                    row_vals.append(f"{col_let}: {f_val} (= {v_val})")
            print(f"  Row {r} [{opt_name}]: {', '.join(row_vals)}")
            
    # 3. Discount Structure
    print("Discount Structure:")
    for r in range(10, 20):
        # check row
        vals = [ws_v.cell(row=r, column=c).value for c in range(1, 10)]
        if any(v == 'Discount Structure' or v in ['P', 'Pipe', 'F', 'Fittings'] for v in vals if v is not None):
            non_empty = [f"{get_column_letter(c)}{r}: {ws_v.cell(row=r, column=c).value}" for c in range(1, 10) if ws_v.cell(row=r, column=c).value is not None]
            print(f"  Row {r}: {', '.join(non_empty)}")

    # 4. Quantity Inputs
    print("Quantity Inputs:")
    for r in range(15, ws_v.max_row + 1):
        for c in range(1, 10):
            val = str(ws_v.cell(row=r, column=c).value)
            if any(q in val.lower() for q in ['toilet', 'kitchen', 'floor', 'flat', 'tower']):
                # next cell might be value
                v_cell = ws_v.cell(row=r, column=c+1).value
                print(f"  Cell {get_column_letter(c)}{r}='{val}', Value in {get_column_letter(c+1)}{r}='{v_cell}'")

    # 5. Product-Cost Ratio or Auxiliary Blocks
    # Check any cells beyond column K or rows beyond 20
    aux_cells = []
    for r in range(1, ws_v.max_row + 1):
        for c in range(1, ws_v.max_column + 1):
            if c > 13 or r > 20:
                fv = ws_f.cell(row=r, column=c).value
                vv = ws_v.cell(row=r, column=c).value
                if fv is not None or vv is not None:
                    aux_cells.append(f"{get_column_letter(c)}{r}: {fv} (val={vv})")
    if aux_cells:
        print(f"Auxiliary / Ratio Cells ({len(aux_cells)} cells):")
        for ac in aux_cells[:15]:
            print(f"  {ac}")
        if len(aux_cells) > 15:
            print(f"  ... and {len(aux_cells)-15} more.")
    else:
        print("Auxiliary / Ratio Cells: NONE")
