from openpyxl import load_workbook

wb = load_workbook('sheets.xlsx')
sheet1 = wb.get_sheet_by_name('sheet1')
sheet2 = wb.get_sheet_by_name('sheet2')
sheet3 = wb.get_sheet_by_name('sheet3')
sheet4 = wb.get_sheet_by_name('sheet4')
sheet5 = wb.get_sheet_by_name('sheet5')
Master = wb.get_sheet_by_name('Master')
