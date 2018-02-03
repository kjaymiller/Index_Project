from openpyxl import load_workbook
wb = load_workbook('Bulk Master Index.xlsx', data_only=True)
sheet = wb.active

for row in sheet:
    values = row[2].value
    sorted_value = sorted(row[2].value.split(', '))
    print(sorted_value)
