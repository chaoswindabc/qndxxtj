from openpyxl import load_workbook


class cl:
    grade = 0
    name = ''
    total = 0  #总计
    unfi = 0  #未完成
    per = 0  #完成率

    def __init__(self,n,t):
        self.name = n
        self.total = t

    def cal(self):
        self.per = 1 - (self.unfi/self.total)
        return self.per


def cnt(nam):
    for i in range(0,sh3.max_row):
        if arr[i].name == nam:
            arr[i].unfi+=1
            return
    return

wb1 = load_workbook("未完成名单.xlsx")
wb2 = load_workbook("总名单.xlsx")
wb3 = load_workbook("总人数表.xlsx")
sh1 = wb1["Sheet1"]
sh2 = wb2["Sheet1"]
sh3 = wb3["Sheet1"]
# total_stu=sh1.max_row
# print(total_stu)

arr=[]
for i in range(1,sh3.max_row+1):
    arr.append(cl(sh3.cell(i,1).value, sh3.cell(i,2).value))

# for i in range(3,total_stu):
#     cell = sh1.cell(i, 2)
#     print(cell.value)

# for col in sh1.rows:
#     print(col[1].value)

for col in sh1.rows:
    cnt(col[1].value)

for i in range(0,sh3.max_row):
    arr[i].cal()
    print(arr[i].name, arr[i].per)