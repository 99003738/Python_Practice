from openpyxl import Workbook

import time

pathList = []
inputList = []
Dict = dict({"RF_TextFilePath": 0, "RF_ManualPath" : 0, "RF_ManualFilterInput": 0, "Flag_showData": 0})
inputDict = {}
outputPath = []

def textfilenameread():
    pass


def textfilepathread():

    textfilepath = input("Enter your .txt Path here:  ")
    file = open(textfilepath, "r")
    context = file.read()
    textpathlist = context.split("\n")
    # print(textpathlist)
    Dict["RF_TextfilePath"] = 1
    for item in textpathlist:
        pathList.append(item)



def terminalpathinput():
    print("Copy all the below and hit double Enter key :")
    while True:  # directly copy all the path in terminal with new line and hit enter and enter
        line = input()  # the directory file will loaded into the list form and can be extract by list index
        if line:
            pathList.append(line)
        else:
            break


def manuentrypath():
    manuInput = input("Copy Your Path Here : ")
    pathList.append(manuInput)
    Dict["RF_ManualPath"] = 1
    userInput = input("Want to Add more:: Y/N : ")
    while True:
        if userInput == 'Y' or userInput == 'y':
            manuInput = input("Copy : ")
            pathList.append(manuInput)
        elif userInput == 'N' or userInput == 'n':
            break
        userInput = input("Want to Add more:: Y/N : ")


def manuentryname():
    pass


def user_selection():

    print("Enter the choice given below. Use Number for your selected choice")
    print("1. Add Text File  Name")
    print("2. Give path of .txt which contain Workbook path ")
    print("3. Copy All the Workbook Path directly to terminal")
    print("4. Manually You want to give the Workbook Path")
    print("5. Manually Giving the Workbook Name Present in the same directory")
    print("6. Exit ")
    choice = int(input("Waiting for your choice : "))
    if choice == 1:
        textfilenameread()
    elif choice == 2:
        textfilepathread()
    elif choice == 3:
        terminalpathinput()
    elif choice == 4:
        manuentrypath()
    elif choice == 5:
        manuentryname()
    elif choice == 6:
        return


def manual_filter_input():

    print("Enter Maximum Three Identity :: Format :: Col_Name,value ::: Ex- xyz,123abc :::: PRESS Enter key")
    filterinput = input("Input Data Here : ")
    # print(filterinput)
    inputlist1 = filterinput.split(",")                              # Here taking input in list but filtering the data according to second data of list.
    inputList.append(inputlist1[1])
    # inputDict[inputlist1[0]] = inputlist1[1]
    Dict["RF_ManualFilterInput"] = 1
    userInput = input("Want to add more filter:: Y/N : ")
    while True:
        if userInput == 'Y' or userInput == 'y':
            filterinput = input("Input Data Here : ")
            inputlist1 = filterinput.split(",")
            inputList.append(inputlist1[1])
            # inputDict[inputlist1[0]] = inputlist1[1]

        elif userInput == 'N' or userInput == 'n':
            break
        userInput = input("Want to add more filter:: Y/N : ")


def show_me_data():
    print("1. Search By input")
    print("2. Automatic Search")
    choice = int(input(" Enter Choice"))

    if choice == 1:
        manual_filter_input()
        Dict["Flag_showData"] = 1
    elif choice ==2:
        Dict["Flag_showData"] = 0





def user_choice_selection():
    print(" 1. Filter the Data according to your choice ")
    print(" 2. Want to search filter input ")
    userChoice = int(input("Enter Your choice here : "))
    if userChoice == 1:
        manual_filter_input()
    elif userChoice == 2:
        show_me_data()


def create_workbook():
    book = Workbook()
    sheet = book.active
    sheet.title = "Sheet1"
    book.save(outputPath[0])


def desktop_output():
    import os
    name = input("Enter New Work Book Name: ")
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    newpath = desktop + "\\" + name + ".xlsx"
    outputPath.append(newpath)
    create_workbook()


def give_path():

    choice = input(" Paste Your Path Here : ")
    outputPath.append(choice)


def outputResult_path():
    print("\n")
    print(" Give Choice For Output Result")
    print("1. Create New WorkBook on Desktop")
    print("2. Give Existing Work Book Path")
    choice = int(input("Enter Your Choice Here : "))
    if choice == 1:
        desktop_output()
    elif choice == 2:
        give_path()

# user_selection()
# user_choice_selection()
# outputResult_path()

# print(pathList)



