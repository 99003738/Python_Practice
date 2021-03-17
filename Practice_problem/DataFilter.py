import UserInput
from openpyxl import load_workbook

ValidatingInputFlag = 0


def sheettraversing(sheet):
    maxrows = sheet.max_row
    for i in range(2, maxrows + 1):
        a1 = sheet.cell(row=i, column=1)
        if UserInput.inputList[0] == a1:
            a2 = sheet.cell(row=i, column=2)
            if UserInput.inputList[1] == a2:
                a3 = sheet.cell(row=i, column=3)
                if UserInput.inputList[2] == a3:
                     return 1


def validating_input():
    print("you are in validating input sect")

    lenofinputpath = len(UserInput.pathList)
    if UserInput.RF_MenualPath == 0 and UserInput.RF_ManualFilterInput == 0:

        print(lenofinputpath)
    else:
        print("Exporting path error")
    for i in range(0, lenofinputpath): # ? check the value here weather iterating correct or not
        path = UserInput.pathList[i]
        workbook = load_workbook(path)
        for sheet in workbook.worksheets:
            value = sheettraversing(sheet)
            if value == 1:
                print("Data matched found and authentic")



UserInput.user_selection()
UserInput.user_choice_selection()
print(UserInput.inputList)
validating_input()