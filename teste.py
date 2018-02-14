from openpyxl import Workbook

string = '   teste  123   '
print(string)
print(string.rstrip())
print(string.strip())

listteste = []
for x in range(10):
        listteste.append(x)

for x in range(len(listteste)):
    if x < len(listteste) and divmod(listteste[x],2)[1]  == 0 :
        listteste.pop(x)

print(listteste)