import sys, os
ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
RAW_BOQ_DIR = os.path.join(ROOT_DIR, "data", "raw_boq")
sys.path.insert(0, os.path.join(ROOT_DIR, "backend"))
import huliot_core, openpyxl

files = [
    'Huliot BOQ Summary AVINEA BY SIGMA 13.08.2026 - 09.10.2026.xlsx',
    'Huliot BOQ Summary BDR Hospital_08.07.2026.xlsx',
    'Huliot BOQ Summary For Empire Developers 22.10.2026 - 31.10.2026.xlsx',
    'Huliot BOQ Summary For Supernova by Kanifnath Project, 13.07.2026.xlsx',
    'Huliot BOQ Summary Tribeca Lalbaug_10.12.2025- 02.11.2026 (1).xlsx',
    'Huliot BOQ Summary_CASA Terazza_04.11.2026.xlsx',
    'Huliot BOQ VTP Monarque_15.10.2025.xlsx',
    'Huliot BOQ For Shubh tristar Project,03.07.2025.xlsx',
    'Huliot BOQ IHP Kalpataru-Sinhagad Road 23.09.2025.xlsx',
    'Huliot BOQ Holystico ASTRANA 11.09.2025.xlsx',
    'Huliot BOQ VTP Earth Tower 1,2,3,4,5,6,7_10.10.2024.xlsx',
    'Huliot BOQ VTP Verve Tower 2,3,4,5,6_16.02.2024.xlsx',
    'Huliot BOQ Kingdom 57 28.06.2025 25.08.2025.xlsx',
    'Huliot BOQ Satva Commercial STD Length 14.10.2025, 11.11.2025.xlsx',
    'Huliot BOQ RSV CANTEEN 26.08.2025.xlsx',
    'Huliot BOQ Aster 108 Punawale Tower A_22.01.2026.xlsx',
    'Huliot BOQ Aster 108 Punawale Tower B_22.01.2026.xlsx',
    'Huliot BOQ Aster 108 Punawale Tower C_17.1.26_22.1.26.xlsx',
    'Huliot BOQ Aster 108 Commercial Area Toilet_31.01.2026.xlsx'
]

print(f"=== RUNNING FULL REGRESSION SUITE ACROSS ALL {len(files)} FILES ===\n")

all_passed = True
for f in files:
    src_f = os.path.join(RAW_BOQ_DIR, f)
    out_f = os.path.join(ROOT_DIR, 'test_reg_' + os.path.basename(f))
    try:
        huliot_core.generate_summary_sheet(src_f, out_f)
        wb = openpyxl.load_workbook(out_f, data_only=False)
        ws = wb['Summary Sheet']
        is_hosp = ('hospital' in f.lower()) or (ws['D19'].value == 0)

        failures = []
        for r in range(5, 8):
            opt = ws.cell(r, 3).value
            if not opt:
                continue
            t_cost = str(ws.cell(r, 5).value or '')
            k_cost = str(ws.cell(r, 7).value or '')
            gt_cost = str(ws.cell(r, 9 if not is_hosp else 7).value or '')

            if t_cost in ['=0', '0', '', 'None']:
                failures.append(f'Row {r} ({opt}): Toilet Cost is empty/zero ({t_cost})')
            if not is_hosp and k_cost in ['=0', '0', '', 'None']:
                failures.append(f'Row {r} ({opt}): Kitchen Cost is empty/zero ({k_cost})')

        toilets = ws['D18'].value
        kitchens = ws['D19'].value
        floors = ws['D20'].value
        towers = ws['D21'].value

        if not toilets or toilets == 0:
            failures.append(f'Toilets count is invalid: {toilets}')
        if not is_hosp and (not kitchens or kitchens == 0):
            failures.append(f'Kitchens count is invalid: {kitchens}')
        if not floors or floors == 0:
            failures.append(f'Floors count is invalid: {floors}')
        if not towers or towers == 0:
            failures.append(f'Towers count is invalid: {towers}')

        # Check Tower table if multi-tower
        if towers > 1:
            if not ws['C24'].value or 'Tower-Wise' not in str(ws['C24'].value):
                failures.append(f'Multi-tower project ({towers} towers) missing Tower Breakdown Table at C24')

        if failures:
            print(f"❌ FAILED: {os.path.basename(f)}")
            for fail in failures:
                print(f"   - {fail}")
            all_passed = False
        else:
            e5 = ws['E5'].value
            g5 = ws['G5'].value
            i5 = ws['I5'].value
            j5 = ws['J5'].value # Floor cost
            print(f"✅ PASSED: {os.path.basename(f)}")
            print(f"   HT Pro -> Toilet: {str(e5)[:40]}... | Total: {i5} | T:{toilets} K:{kitchens} Flr:{floors} Twr:{towers}")
            if towers > 1:
                print(f"   [Multi-Tower Table Created: C24='{ws['C24'].value}' with {towers} towers]\n")
            else:
                print()
    except Exception as e:
        print(f"❌ EXCEPTION in {os.path.basename(f)}: {e}\n")
        all_passed = False
    finally:
        if os.path.exists(out_f):
            os.remove(out_f)

print(f"=== FINAL RESULT: {f'ALL {len(files)} FILES PASSED 100%' if all_passed else 'FAILURES FOUND'} ===")
