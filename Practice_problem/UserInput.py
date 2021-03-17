
pathList = []
inputList = []
RF_TextFilePath = 0
RF_MenualPath = 0
inputDict = {}
RF_ManualFilterInput = 0


def textfilenameread():
    pass


def textfilepathread():

    textfilepath = input("Enter your .txt Path here:  ")
    file = open(textfilepath, "r")
    context = file.read()
    pathList = context.split("\n")
    RF_TextFilePath = 1


def terminalpathinput():
    pass


def manuentrypath():
    manuInput = input("Copy Your Path Here : ")
    pathList.append(manuInput)
    RF_MenualPath = 1
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
    print("1. Add Workbook path from .txt file with name")
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
    #print(filterinput)
    inputlist1 = filterinput.split(",")
    inputList.append(inputlist1[1])
    #inputDict[inputlist1[0]] = inputlist1[1]
    RF_ManualFilterInput = 1
    userInput = input("Want to add more filter:: Y/N : ")
    while True:
        if userInput == 'Y' or userInput == 'y':
            filterinput = input("Input Data Here : ")
            inputlist1 = filterinput.split(",")
            inputList.append(inputlist1[1])
            # inputDict[inputlist1[0]] = inputlist1[1]

        elif  userInput == 'N' or userInput == 'n':
            break
        userInput = input("Want to add more filter:: Y/N : ")


def show_column_header():
    pass


def user_choice_selection():
    print(" 1. Filter the Data according to your choice ")
    print(" 2. Want to see all the Column Header name ")
    userChoice = int(input("Enter Your choice here : "))
    if userChoice == 1:
        manual_filter_input()
    elif userChoice == 2:
        show_column_header()
# user_selection()
# print(pathList)
# user_choice_selection()
# print(inputDict)