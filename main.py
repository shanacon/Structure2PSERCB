from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.styles import Alignment

InitList_en = ['name', 'type', 'Bc/Dc', 'Hc', 'Los(%)', 'No1', 'Num1', 'No2', 'Num2', 'h1', 'No', 'Num', 'S', 'Nci', 'whichFloor']
InitList_zh= ['斷面名稱', '型式', '寬度/直徑', '深度', '主筋鋼筋比', '主筋號數(1)', '主筋根數(1)', '主筋號數(2)', '主筋根數(2)', 
              '一樓柱淨高', '橫向箍、繫筋號數', '橫向箍、繫筋根數', '箍筋間距', '柱根數', '所在樓層']
class Case :
    def __init__(self, name, type, BC, HC, No1, Num1, No2, Num2, H1, No, Numx, Numy, S, Nci, whichFloor):
        self.name = name  
        self.type = type  
        self.BC = BC  
        self.HC = HC  
        self.No1 = No1  
        self.Num1 = Num1  
        self.No2 = No2  
        self.Num2 = Num2  
        self.H1 = H1  
        self.No = No  
        self.Numx = Numx
        self.Numy = Numy
        self.S = S  
        self.Nci = Nci  
        self.whichFloor = whichFloor
Err = False
FHdic = {}
CNdic = {}
ALLCASE = []
## load CNK
CNK1 = open('CNK1.INP', 'r')
## Load useless data
KeepRead  = True
while KeepRead :
    line = CNK1.readline()
    if line.find('col.data') != -1:
        KeepRead = False
C = CNK1.readline().split()[0]
for i in range(int(C)):
    line = CNK1.readline().split(',')
    CNdic[line[1]] = int(line[2])
## load MASS103
MASS = open('MASS103.INP', 'r')
## Load useless data
KeepRead  = True
while KeepRead :
    line = MASS.readline()
    if line.find('FLOOR') != -1:
        KeepRead = False
## load CXX
CXX = open('CXX.DAT', 'r')
## get Floor and Cross section
F = CXX.readline().split()[1]
## Load useless data in CXX
for i in range(int(C)):
    CXX.readline()
for i in range(int(F)):
    CXX.readline()
for i in range(int(F)):
    CXX.readline()
## get h1 from MASS
tmp = "{:.1f}".format(float(MASS.readline().split()[1]) * 100.0)
for i in range(int(F)) :
    line = MASS.readline().split()
    if line[0][-1] == 'L':
        fin = line[0][:-1]
    elif line[0][-1] == 'A':
        fin = line[0] + 'F'
    else:
        fin = line[0]
    FHdic[fin] = tmp
    tmp = "{:.1f}".format(float(line[1]) * 100.0)
CXX_Data = CXX.readlines()
CaseCXX = 0
while CaseCXX < len(CXX_Data):
    #floor and cross
    floor = CXX_Data[CaseCXX].split()[0] + 'F'
    cross = CXX_Data[CaseCXX].split()[2]
    # name
    name = floor + cross
    ## BC HC type
    BC = CXX_Data[CaseCXX + 1].split()[0]
    HC = CXX_Data[CaseCXX + 1].split()[1]
    if HC == '0':
        type = 'CIRL'
    else :
        type = 'RECT'
    ## No1 No2
    No1 = CXX_Data[CaseCXX + 2].split()[0]
    if CXX_Data[CaseCXX + 2].split()[1] == '0':
        No2 = str(int(No1) - 1)
    else :
        No2 = CXX_Data[CaseCXX + 2].split()[1]
    No1 = '#' + No1
    No2 = '#' + No2
    ## Num1 Num2
    Num1 = int(CXX_Data[CaseCXX + 3].split()[0]) + int(CXX_Data[CaseCXX + 4].split()[0]) - 4
    if Num1 < 0:
        Num1 = 0
    Num2 = 0
    H1 = FHdic[floor]
    if CXX_Data[CaseCXX + 5].split()[0] != CXX_Data[CaseCXX + 5].split()[3]:
        Err = True
        print('error in No')
        break
    No = '#' + CXX_Data[CaseCXX + 5].split()[0]
    Numx = int(CXX_Data[CaseCXX + 4].split()[3]) + 2
    Numy = int(CXX_Data[CaseCXX + 3].split()[3]) + 2
    S = CXX_Data[CaseCXX + 5].split()[1]
    Nci = CNdic[cross]
    whichFloor = floor
    ALLCASE.append(Case(name, type, BC, HC, No1, Num1, No2, Num2, H1, No, Numx, Numy, S, Nci, whichFloor))
    CaseCXX = CaseCXX + 6
# initial template excel
NewWb = Workbook()
sheetX = NewWb.active
sheetX.title = '一般柱-X'
for i in range(65,80):
    sheetX.column_dimensions[chr(i)].width = 15.0
SheetRange = sheetX['A1':'O1']
i = 0
for item in SheetRange[0]:
    item.value = InitList_en[i]
    item.font = Font(name = 'Times New Roman', bold=True)
    item.alignment = Alignment(horizontal = 'center')
    i = i + 1
SheetRange = sheetX['A2':'O2']
i = 0
for item in SheetRange[0]:
    item.value = InitList_zh[i]
    item.font = Font(name = 'Times New Roman', bold=True)
    item.alignment = Alignment(horizontal = 'center')
    i = i + 1
# ## Final Adjustment
# for row in range(2, sheetX.max_row):
#     for column in range(sheetX.max_column):
#         sheetX.cell(row=row+1,column=column+1).alignment = Alignment(horizontal = 'center')
#         sheetX.cell(row=row+1,column=column+1).font = Font(name = 'Times New Roman')
# NewWb.save('test.xlsx')