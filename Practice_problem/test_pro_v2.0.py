import openpyxl








if __name__ == '__main__':
    my_wb = openpyxl.Workbook()
    my_sheet = my_wb.active

# my_wb.create_sheet(index = 2 , title = "Master_sheet")
# my_wb.save("D:\Git_Python\Practice_problem\Book1.xlsx")

    my_path = "C:/Users/Shivam/PycharmProjects/Python_Practice/Practice_problem/Book1.xlsx"
    my_wb_obj = openpyxl.load_workbook(my_path)
    my_sheet_obj = my_wb_obj.active
    max_rows = my_sheet_obj.max_row
    print(max_rows)
    my_column = my_sheet_obj.max_column
    print(my_column)

    print("Enter the PS Number")
    ps = int(input())
    print("Enter the name : ")
    name = input()
    print("Enter the email Id")
    email = input()
    psFlag = 0
    nameFlag = 0
    emailFlag = 0
    for i in range(1, max_rows + 1):
        cell_obj = my_sheet_obj.cell(row=i, column=1)
        if ps == cell_obj.value:
            print(cell_obj.value)
            psFlag = 1
            cell_obj = my_sheet_obj.cell(row=i, column=2)
            if cell_obj.value == name:
                print(cell_obj.value)
                nameFlag = 1
                cell_obj = my_sheet_obj.cell(row=i, column=3)
                if cell_obj.value == email:
                    print(cell_obj.value)
                    emailFlag = 1
        if psFlag == 1 and nameFlag == 1 and emailFlag == 1:
            print("Match found in records")
            #printRecords()


def printRecords():
    my_wb.create_sheet(index=1, title='Master')
    my_wb.save("C:/Users/Shivam/PycharmProjects/Python_Practice/Practice_problem/Book1.xlsx")


