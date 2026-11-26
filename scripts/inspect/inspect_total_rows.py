import openpyxl

files = [
    ('File 1: Avinea', 'Huliot BOQ Summary AVINEA BY SIGMA 13.08.2026 - 09.10.2026.xlsx', [
        ('Typical Floor BOQ  (T)', 63, ['AP', 'AW', 'BD']),
        ('Typical Floor BOQ  (K)', 63, ['V', 'AC', 'AJ']),
        ('Typical Floor BOQ ', 63, ['BB', 'BI', 'BP'])
    ]),
    ('File 2: BDR Hospital', 'Huliot BOQ Summary BDR Hospital_08.07.2026.xlsx', [
        ('Typical Floor BOQ G to 7', 83, ['HW', 'ID']),
        ('Diversion BOQ 4 Floor ', 55, ['CG', 'CN'])
    ]),
    ('File 3: Empire', 'Huliot BOQ Summary For Empire Developers 22.10.2026 - 31.10.2026.xlsx', [
        ('Typical Floor BOQ  (T)', 95, ['AH', 'AW', 'BB']),
        ('Typical Floor BOQ  (K)', 95, ['Y', 'AN', 'AS'])
    ]),
    ('File 4: Supernova', 'Huliot BOQ Summary For Supernova by Kanifnath Project, 13.07.2026.xlsx', [
        ('Typical Floor BOQ  (T)', 75, ['AJ', 'AQ', 'AX']),
        ('Typical Floor BOQ  (K)', 75, ['R', 'Y', 'AF'])
    ]),
    ('File 5: Tribeca', 'Huliot BOQ Summary Tribeca Lalbaug_10.12.2025- 02.11.2026 (1).xlsx', [
        ('Typical Floor BOQ  (T)', 75, ['X', 'AE', 'AL']),
        ('Typical Floor BOQ  (K)', 75, ['N', 'U', 'AB'])
    ]),
    ('File 6: CASA Terazza', 'Huliot BOQ Summary_CASA Terazza_04.11.2026.xlsx', [
        ('Typical Floor BOQ  (T)', 59, ['DU', 'EB', 'EI']),
        ('Typical Floor BOQ  (K)', 59, ['BS', 'BZ', 'CG']),
        ('Diversion BOQ ', 41, ['AB', 'AI'])
    ])
]

for label, fname, sheets_info in files:
    print(f"\n=======================================================")
    print(f"{label}")
    wb_f = openpyxl.load_workbook(fname, data_only=False)
    wb_v = openpyxl.load_workbook(fname, data_only=True)
    for sname, r_idx, cols in sheets_info:
        ws_f = wb_f[sname]
        ws_v = wb_v[sname]
        print(f"\nSheet '{sname}' Row {r_idx}:")
        print(f"  Col A: '{ws_v.cell(row=r_idx, column=1).value}'")
        print(f"  Col B: '{ws_v.cell(row=r_idx, column=2).value}'")
        print(f"  Col C: '{ws_v.cell(row=r_idx, column=3).value}'")
        # Check rows around r_idx
        for check_r in [r_idx - 1, r_idx, r_idx + 1]:
            if 1 <= check_r <= ws_v.max_row:
                print(f"  Row {check_r} A-C: {[ws_v.cell(row=check_r, column=c).value for c in range(1, 4)]}")
        for c in cols:
            f_val = ws_f[f"{c}{r_idx}"].value
            v_val = ws_v[f"{c}{r_idx}"].value
            # Check header at row 12, 13, 14
            col_idx = openpyxl.utils.column_index_from_string(c)
            # Find option header by looking at row 11, 12, 13 in this col or merged
            opt_hdr = None
            for r_hdr in [11, 12, 13]:
                hv = ws_v.cell(row=r_hdr, column=col_idx).value
                if not hv:
                    # check merged cells
                    for m in ws_v.merged_cells.ranges:
                        if m.min_row <= r_hdr <= m.max_row and m.min_col <= col_idx <= m.max_col:
                            hv = ws_v.cell(row=m.min_row, column=m.min_col).value
                            break
                if hv:
                    opt_hdr = f"R{r_hdr}:{hv}"
                    break
            sub_hdr = ws_v.cell(row=14, column=col_idx).value
            print(f"  Col {c}{r_idx}: OptHeader={opt_hdr} | SubHeader(R14)={sub_hdr} | Formula={f_val} | Value={v_val}")
