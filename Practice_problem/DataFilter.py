import UserInput
from openpyxl import load_workbook

ValidatingInputFlag = 0


def sheettraversing(sheet):
    print("you are in sheettraversing")
    maxrows = sheet.max_row
    maxcolumn = sheet.max_column
    print(maxrows)
    print(maxcolumn)
    loi_list = len(UserInput.inputList)
    for listdata in range(0, loi_list):
        for i in range(2, maxrows):                                   # i is rows and j is column
            for j in range(1,3):
                data_at_cell1 = sheet.cell(row= i, column= j)
                if data_at_cell1.value == UserInput.inputList[listdata]:
                    data_at_cell1 = sheet.cell(row= i, column = 1)
                    data_at_cell2 = sheet.cell(row= i, column= 2)
                    data_at_cell3 = sheet.cell(row= i, column= 3)
                    v1 = data_at_cell1.value
                    v2 = data_at_cell2.value
                    v3 = data_at_cell3.value
                    v = [v1,v2,v3]
                    print(v)
                    for check in range(0, 3):
                        for inner in range(0,loi_list):
                            if v[check] == UserInput.inputList[inner]:
                                v[check] = 1
                        print(v[check])
                    if (v[0] == 1 and v[1] == 1) and v[2] == 1:
                        print("Input matched found")

                    # for checklist in range(0, loi_list):
                    #
                    #     checklist = checklist + 1
                    #     if data_at_cell2.value == UserInput.inputList[checklist] or data_at_cell3.value == UserInput.inputList[checklist] or data_at_cell1.value == UserInput.inputList[checklist]:
                    #         checklist = checklist+1
                    #         if data_at_cell3.value == UserInput.inputList[checklist] or data_at_cell3.value == UserInput.inputList[checklist] or data_at_cell3.value == UserInput.inputList[checklist]:
                    #             print("Input matched found")
                    #             return 1
                    #         else:
                    #             print("Do some thing else")



def validating_input():
    print("you are in validating input sect")
    lenofinputpath = len(UserInput.pathList)

    if UserInput.Dict["RF_ManualPath"] ==1 and UserInput.Dict["RF_ManualFilterInput"] ==1:
        print(lenofinputpath)
    else:
        print("Exporting path error")
    for i in range(0, lenofinputpath): # ? check the value here weather iterating correct or not
        count = 0
        path = UserInput.pathList[i]
        workbook = load_workbook(path)
        sheetlist = workbook.sheetnames
        print(sheetlist)
        totalsheet = len(sheetlist)
        print(totalsheet)
        for sheet in range(0, totalsheet):
            # sheets = workbook[sheetlist[sheet]]
            sheettraversing(workbook[sheetlist[sheet]])
            count = count+1
        print("Back")
        print(count)

UserInput.user_selection()
UserInput.user_choice_selection()
print(UserInput.inputList)
validating_input()