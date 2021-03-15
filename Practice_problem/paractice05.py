n = int(input())
sum = 0
var = []
if n>0:
    print(n)
    for i in range(0,n):
        var.append(int(input()))
    print(var)
    for i in range(0,n):
        sum = sum + var[i]
print(sum)