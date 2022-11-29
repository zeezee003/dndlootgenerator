import openpyxl as xl


wb = xl.load_workbook(filename='dndloot.xlsx', data_only=True)
ws = wb['Loot']
row_range = ws['M42'].value
gemArray = []
d12Array = []

for col in ws.iter_cols(min_row=4,min_col=14,max_row=(3+row_range),max_col=14):
    for cell in col:
        gemArray.append(cell.value)
        #print(cell.value)

print(f"This is the gem array {gemArray}")

for col in ws.iter_cols(min_row=4,min_col=15,max_row=(4+11),max_col=15):
    for cell in col:
        d12Array.append(cell.value)

print(f"This is the d12 array {d12Array}")
print(f"There are {row_range} gems in this haul")
j = 0
for i in gemArray:
    print(f"There are {gemArray.count(i)} {i}'s")