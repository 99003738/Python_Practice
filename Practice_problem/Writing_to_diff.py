from openpyxl import load_workbook

wb = load_workbook('sheets.xlsx')
sheet1 = wb.get_sheet_by_name('January')
sheet2 = wb.get_sheet_by_name('February')
sheet3 = wb.get_sheet_by_name('March')

maxRows = sheet1.max_row
maxColumn = sheet1.max_column
print("Enter your PS Number :")
ps = int(input())
print("Enter your Name : ")
name = input()
print("Enter your Email : ")
email = (input())

for i in range(1, maxRows+1):
    a1 = sheet1.cell(row=i, column=1)
    if a1.value == ps:
        a2 = sheet1.cell(row=i, column=2)
        if a2.value == name:
            a3 = sheet1.cell(row=i, column=3)
            if a3.value == email:
                flag = 1
                print(a1.value, a2.value, a3.value)
k = 2
l = 1

for i in range(1, maxRows+1):
    a1 = sheet1.cell(row=i, column=1)
    a2 = sheet1.cell(row=i, column=2)
    a3 = sheet1.cell(row=i, column=3)
    if (a1.value == ps and a3.value == email) and a2.value == name:
        k = k + 1
        for j in range(4,maxColumn+1):
            sheet2.cell(row=k, column=l).value = sheet1.cell(row=i, column=j).value
            l = l+1

    l = 1
for i in range(1, maxRows+1):
    a1 = sheet3.cell(row=i, column=1)
    a2 = sheet3.cell(row=i, column=2)
    a3 = sheet3.cell(row=i, column=3)
    if (a1.value == ps and a3.value == email) and a2.value == name:
        k = k + 1
        for j in range(4,maxColumn+1):
            sheet2.cell(row=k, column=l).value = sheet3.cell(row=i, column=j).value
            l = l+1

    l = 1
wb.save('sheets.xlsx')


# coding to copy whole data
# for i in range(1, maxRows+1):
#     for j in range(1, sheet1.max_column+1):
#         sheet2.cell(row=i , column= j).value = sheet1.cell(row=i, column=j).value
# wb.save('sheets.xlsx')

