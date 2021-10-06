import openpyxl as xl

wb = xl.load_workbook('file.xlsx')

sheet = wb['Sheet1']
cell = sheet['A1']
print(cell.value)
cell = sheet.cell(1, 1)
print(cell.value)
print(sheet.max_row)

for row in range(2, sheet.max_row + 1):
    cell = sheet.cell(row, 3)
    c1 = 0.9 * cell.value
    c2 = sheet.cell(row,4)
    c2.value = c1

wb.save('file2.xlsx')

