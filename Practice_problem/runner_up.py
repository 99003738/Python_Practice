

n = int(input())
arr = list(map(int, input().rstrip().split()))

for i in range(0, n-1):
    arr.append(int(input()))
first = arr[0]
second = arr[1]
j = 1
if n >= 2 and n <= 10:
    for i in range(1, n):
        if arr[i] > first:
            second = first
            first = arr[i]
        else:
            second = arr[i]
print(first)
print(second)
