import openpyxl as xl


intputlist =[]
entry = int(input("Enter a integer value : "))

intputlist.append(entry)
name = int(input("enter an string value : "))
intputlist.append(name)
stringconversio = input("Enter a string value : ")
intputlist.append(stringconversio)


print(type(intputlist[0]))
print(intputlist[0])
print(type(intputlist[1]))
print(intputlist[1])
print(type(intputlist[2]))
print(intputlist[2])

# con = str(intputlist[0])
# con1 = str(intputlist[1])
# con2 = str(intputlist[2])
# print(type(con))
# print(type(con1))
# print(type(con2))
# print("\n")

if str(entry) == str(name):
    print("1")
if str(name) == stringconversio:
    print("1")