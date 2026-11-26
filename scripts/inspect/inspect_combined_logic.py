import openpyxl

wb1 = openpyxl.load_workbook('Huliot BOQ Summary AVINEA BY SIGMA 13.08.2026 - 09.10.2026.xlsx', data_only=False)
ws_t = wb1['Typical Floor BOQ  (T)']
ws_k = wb1['Typical Floor BOQ  (K)']
ws_comb = wb1['Typical Floor BOQ ']

print("File 1: Checking Combined Sheet vs T and K")
# Check cell formulas in Combined sheet row 15
for col in ['A', 'B', 'C', 'D', 'AW', 'AY', 'AZ', 'BA', 'BB', 'BI', 'BP', 'BS']:
    val = ws_comb[f"{col}15"].value
    print(f"  Comb {col}15: {val}")

# Check Total row in Combined sheet: Row 63
print(f"  Comb BB63: {ws_comb['BB63'].value}")
print(f"  Comb BI63: {ws_comb['BI63'].value}")
print(f"  Comb BP63: {ws_comb['BP63'].value}")

# Check T + K:
# In Summary Sheet:
# Row 5 (HT Pro): E5 = T!AW63 (=4935579.96), G5 = K!AC63 (=1181718.33), I5 = E5+G5 (=6117298.29)
# Let's see: BI63 in Combined is 6117298.29!
print("Is BI63 equal to E5+G5?", ws_comb['BI63'].value, "vs", 4935579.96 + 1181718.33)
