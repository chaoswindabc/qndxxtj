from openpyxl import load_workbook,Workbook
from openpyxl.styles import Font,Alignment

class cl:
    grade = 0
    name = ''
    total = 0  #总计
    unfi = 0  #未完成
    per = 0  #完成率

    def __init__(self,n,t):
        self.name = n
        self.total = t

    def __lt__(self,other):
        return self.per < other.per

    def calPercent(self):
        self.per = 1 - (self.unfi/self.total)

    # def getPercent(self):
    #     return self.per

    def addUnfinished(self):
        self.unfi += 1

    # def getClassName(self):
    #     return self.name


def cnt(nam):
    for i in range(0,total_class):
        if arr[i].name == nam:
            arr[i].addUnfinished()
            return
    return

wb1 = load_workbook("未完成名单.xlsx")
# wb2 = load_workbook("总名单.xlsx")
wb3 = load_workbook("总人数表.xlsx")
sh1 = wb1["Sheet1"]
# sh2 = wb2["Sheet1"]
sh3 = wb3["Sheet1"]
total_class=sh3.max_row
# total_stu=sh1.max_row

arr=[]
for i in range(1,total_class+1):
    arr.append(cl(sh3.cell(i,1).value, sh3.cell(i,2).value))

for col in sh1.rows:
    cnt(col[1].value)

for i in range(0,total_class):
    arr[i].calPercent()

arr.sort(reverse=1)

for i in range(0,total_class):
    print(arr[i].name, arr[i].per)

tarGrade=19
year=2022
issue=0

wb4 = Workbook()
sh4 = wb4.active
if issue == 0:
    wbtitle = sh4.cell(1,1).value = str(tarGrade)+'级青年大学习'+str(year)+'年特辑完成率'
else:
    wbtitle = sh4.cell(1,1).value = str(tarGrade)+'级青年大学习'+str(year)+'年第'+str(issue)+'期完成率'
sh4.merge_cells('A1:B1')
sh4.cell(2,1).value = '班级'
sh4.cell(2,2).value = '完成率'

for i in range(0,total_class):
    sh4.cell(i+3,1).value = arr[i].name
    sh4.cell(i+3,2).value = arr[i].per
    # if arr[i].per < 0.7:
    #     sh4.cell(i+3,1).font = Font(color="ff0000")
    #     sh4.cell(i+3,2).font = Font(color="ff0000")

style_1 = Font(u'宋体', size=14) 
style_2 = Alignment(horizontal='center', vertical='center')
for col in sh4["A:B"]:
    for cell in col:
        cell.font = style_1
        cell.alignment = style_2

for cell in sh4["B"]:
    cell.number_format = '0.00%'

sh4.column_dimensions['A'].width = 42.0
sh4.column_dimensions['B'].width = 15.0

for i in range(3,total_class+3):
    if sh4.cell(i,2).value < 0.7:
        for j in range(i,total_class+3):
            sh4.cell(j,1).font += Font(color="ff0000")
            sh4.cell(j,2).font += Font(color="ff0000")
        break

wb4.save(wbtitle+".xlsx")
