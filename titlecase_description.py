from index_project import wb, sheets, titlecase_description

for sheet in sheets:
    sheet = wb[sheet]
    for index, row in enumerate(sheet.rows):
        titlecase_description(sheet, (index, row), column=2)

wb.save('titlecase_description.xlsx')
