import openpyxl

wb = openpyxl.load_workbook('Huliot BOQ Summary AVINEA BY SIGMA 13.08.2026 - 09.10.2026.xlsx', data_only=True)
ws_comb = wb['Typical Floor BOQ ']

print("File 1: Typical Floor BOQ  (Combined)")
print("Col BS row 14-20:")
for r in range(13, 25):
    print(f"Row {r:2d} -> BS: '{ws_comb.cell(row=r, column=openpyxl.utils.column_index_from_string('BS')).value}' | BB: {ws_comb.cell(row=r, column=openpyxl.utils.column_index_from_string('BB')).value} | BI: {ws_comb.cell(row=r, column=openpyxl.utils.column_index_from_string('BI')).value} | BP: {ws_comb.cell(row=r, column=openpyxl.utils.column_index_from_string('BP')).value} | B: '{ws_comb.cell(row=r, column=2).value}'")

# Let's check unique values in BS15:BS62
bs_vals = set(ws_comb.cell(row=r, column=openpyxl.utils.column_index_from_string('BS')).value for r in range(15, 63))
print("Unique BS values:", bs_vals)
