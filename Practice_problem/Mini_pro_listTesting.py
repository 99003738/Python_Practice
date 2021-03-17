from openpyxl import load_workbook

pathList = []

ps_number = 0

# Taking the user input for how many excell book  it want to be add.     status OK
def user_input_path():

    user_input = int(input("Enter the number of workbook to add for searching :"))
    for i in range(0, user_input):
        pathList.append(input("Copy here path i+1 :"))
    length_list = len(pathList)
    for i in pathList:
        print(i, end=" ")   # status OK

def search_input()
    global ps_number = int(input("Enter your ps number : "))

