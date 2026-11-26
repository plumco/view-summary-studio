import openpyxl
from openpyxl.utils import get_column_letter

files = [
    ('File 1: Avinea', 'Huliot BOQ Summary AVINEA BY SIGMA 13.08.2026 - 09.10.2026.xlsx', 'Typical Floor BOQ  (T)'),
    ('File 2: BDR Hospital', 'Huliot BOQ Summary BDR Hospital_08.07.2026.xlsx', 'Typical Floor BOQ G to 7'),
    ('File 3: Empire', 'Huliot BOQ Summary For Empire Developers 22.10.2026 - 31.10.2026.xlsx', 'Typical Floor BOQ  (T)'),
    ('File 4: Supernova', 'Huliot BOQ Summary For Supernova by Kanifnath Project, 13.07.2026.xlsx', 'Typical Floor BOQ  (T)'),
    ('File 5: Tribeca', 'Huliot BOQ Summary Tribeca Lalbaug_10.12.2025- 02.11.2026 (1).xlsx', 'Typical Floor BOQ  (T)'),
    ('File 6: CASA Terazza', 'Huliot BOQ Summary_CASA Terazza_04.11.2026.xlsx', 'Typical Floor BOQ  (T)')
]

for label, fname, sname in files:
    print(f"\n=======================================================")
    print(f"{label}: Sheet '{sname}'")
    wb_f = openpyxl.load_workbook(fname, data_only=False)
    ws = wb_f[sname]
    
    # Inspect rows 14 and 15 for pricing blocks
    # Look for headers containing 'SKU', 'Unit Price', 'Discount', 'Total'
    pricing_cols = []
    for c in range(1, ws.max_column + 1):
        v14 = str(ws.cell(row=14, column=c).value)
        if any(h in v14 for h in ['SKU', 'Unit Price', 'Discount', 'Total']):
            pricing_cols.append((c, get_column_letter(c), v14))
            
    print(f"Pricing columns found: {len(pricing_cols)}")
    # Print sample formulas at row 15, 16, 17 for discount and total
    for c, let, hdr in pricing_cols:
        val15 = ws.cell(row=15, column=c).value
        print(f"  Col {let} (R14='{hdr}'): Row 15 = {val15}")
