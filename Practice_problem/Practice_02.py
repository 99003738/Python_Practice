num = int(input())

if (num >= 1) and (num <= 100):
    if num%2 != 0:
        print("Weird")
    elif (num >= 2 and num <= 5) and (num % 2 == 0):
        print("Not Weird")
    elif (num >= 6 and num <= 20) and (num % 2 == 0):
        print("Weird")
    elif num>20 and  num %2 == 0:
        print("Not Weird")