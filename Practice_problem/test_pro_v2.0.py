import openpyxl

my_wb = openpyxl.Workbook()
my_sheet = my_wb.active

#my_wb.create_sheet(index = 2 , title = "Master_sheet")
#my_wb.save("D:\Git_Python\Practice_problem\Book1.xlsx")

val = int(input())
my_path = "D:\Git_Python\Practice_problem\Book1.xlsx"
my_wb_obj = openpyxl.load_workbook(my_path)
my_sheet_obj = my_wb_obj.active

max_rows = my_sheet_obj.max_row
print(max_rows)
my_column = my_sheet_obj.max_column
print(my_column)

for i in range(1 , max_rows+1):

    cell_obj = my_sheet_obj.cell(row = i, column = 1)
    if val == cell_obj.value:
        print(cell_obj.value)
        for j in range(1,my_column+1):
            cell_obj = my_sheet_obj.cell(row = i, column = j)
            print(cell_obj.value)
