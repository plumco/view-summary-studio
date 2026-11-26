import openpyxl

wb = openpyxl.load_workbook('Huliot BOQ Summary AVINEA BY SIGMA 13.08.2026 - 09.10.2026.xlsx', data_only=False)
ws = wb['Summary Sheet']

print("Summary Sheet Styles - File 1:")
for r in [2, 3, 4, 5, 11, 12, 18, 23, 24, 25]:
    for c in ['C', 'D', 'E', 'F', 'I', 'J', 'K']:
        cell = ws[f"{c}{r}"]
        if cell.value is not None:
            font = cell.font
            fill = cell.fill
            border = cell.border
            align = cell.alignment
            nf = cell.number_format
            fg_color = fill.fgColor.rgb if fill and fill.fgColor else None
            f_color = font.color.rgb if font and font.color else None
            print(f"  {c}{r}: Val={cell.value} | Font={font.name}, sz={font.size}, b={font.bold}, col={f_color} | Fill={fill.fill_type}, fg={fg_color} | Align={align.horizontal}/{align.vertical} | NumFmt={nf}")
