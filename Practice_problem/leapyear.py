def is_leap(year):
    if year >= 1900 and year <= 1000000:
        temp = year % 100
        print(temp)
        if year % 4 == 0 and year % 100 != 0 or year % 400 == 0:
            leap = True
        if temp % 4 == 0:
            leap = True
        else:
            leap = False
    else:
        return 0
    return leap


year = int(input())
print(is_leap(year))
