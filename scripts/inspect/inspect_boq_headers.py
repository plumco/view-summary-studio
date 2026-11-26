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
    print(f"\n=======================================================")
    print(f"{label}: {fname}")
    print(f"=======================================================")
    wb = openpyxl.load_workbook(fname, data_only=True)
    for sname in wb.sheetnames:
        if 'Summary' in sname or 'Design' in sname:
            continue
        ws = wb[sname]
        print(f"\n--- Sheet: '{sname}' ---")
        # Find where headers are located (look at rows 1 to 15)
        # Let's print rows 10 to 14 for columns that have values
        for r in range(1, 15):
            non_empty = []
            for c in range(1, min(ws.max_column+1, 150)):
                v = ws.cell(row=r, column=c).value
                if v is not None and str(v).strip():
                    non_empty.append(f"{get_column_letter(c)}{r}: '{str(v).strip()}'")
            if non_empty:
                print(f"  Row {r:2d} (first 8):", non_empty[:8])
