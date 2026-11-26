import openpyxl
import os
import glob

files = [
    'Huliot BOQ Summary AVINEA BY SIGMA 13.08.2026 - 09.10.2026.xlsx',
    'Huliot BOQ Summary BDR Hospital_08.07.2026.xlsx',
    'Huliot BOQ Summary For Empire Developers 22.10.2026 - 31.10.2026.xlsx',
    'Huliot BOQ Summary For Supernova by Kanifnath Project, 13.07.2026.xlsx',
    'Huliot BOQ Summary Tribeca Lalbaug_10.12.2025- 02.11.2026 (1).xlsx',
    'Huliot BOQ Summary_CASA Terazza_04.11.2026.xlsx'
]

for idx, fname in enumerate(files, 1):
    print(f"\n==================================================")
    print(f"FILE {idx}: {fname}")
    print(f"==================================================")
    try:
        wb = openpyxl.load_workbook(fname, data_only=False)
        print("Sheets:", wb.sheetnames)
        for sname in wb.sheetnames:
            ws = wb[sname]
            print(f"  Sheet: '{sname}' | max_row={ws.max_row}, max_column={ws.max_column}, state={ws.sheet_state}")
    except Exception as e:
        print(f"Error opening {fname}: {e}")
