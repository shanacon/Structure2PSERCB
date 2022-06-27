from time import time
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.styles import Alignment
import time
import os
import sys
def myexcepthook(type, value, traceback, oldhook=sys.excepthook):
    oldhook(type, value, traceback)
    input("Press RETURN. ")
sys.excepthook = myexcepthook
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
Num2ExpList = []
print('Makeing dictionary from CNK1 and MASS103...')
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
    elif line[0][-1] == 'F':
        fin = line[0]
    else:
       fin = line[0] + 'F'
    FHdic[fin] = tmp
    tmp = "{:.1f}".format(float(line[1]) * 100.0)
print('Makeing dictionary complete')
print('Loading data from CXX...')
CXX_Data = CXX.readlines()
LineLen = len(CXX_Data)
Progress = 0
CaseCXX = 0
while CaseCXX < LineLen:
    ## set lines
    lines = []
    lines.append(CXX_Data[CaseCXX].split())
    lines.append(CXX_Data[CaseCXX + 1].split())
    lines.append(CXX_Data[CaseCXX + 2].split())
    lines.append(CXX_Data[CaseCXX + 3].split())
    lines.append(CXX_Data[CaseCXX + 4].split())
    lines.append(CXX_Data[CaseCXX + 5].split())
    #floor and cross
    floor = lines[0][0] + 'F'
    cross = lines[0][2]
    # name
    name = floor + cross
    ## BC HC type
    BC = float(lines[1][0])
    HC = float(lines[1][1])
    if BC == 0 and HC == 0:
        CaseCXX = CaseCXX + 6
        continue
    if HC == 0:
        type = 'CIRL'
    else :
        type = 'RECT'
    ## No1 No2
    No1 = lines[2][0]
    if lines[2][1] == '0':
        No2 = str(int(No1) - 1)
    else :
        No2 = lines[2][1]
    No1 = '#' + No1
    No2 = '#' + No2
    ## Num1 Num2
    Num1 = int(lines[3][0]) + int(lines[4][0]) - 4
    if Num1 < 0:
        Num1 = 0
    ## Exception of Num2
    if lines[3][1] != '0' or lines[3][2] != '0' or lines[4][1] != '0' or lines[4][2] != '0':
        Num2ExpList.append(name)
    Num2 = 0
    H1 = float(FHdic[floor])
    if lines[5][0] != lines[5][3]:
        Err = True
        print('error in No. First Element diff with last Element in line 5')
    No = '#' + lines[5][0]
    Numx = int(lines[4][3]) + 2
    Numy = int(lines[3][3]) + 2
    S = int(lines[5][1])
    Nci = CNdic[cross]
    whichFloor = floor
    ALLCASE.append(Case(name, type, BC, HC, No1, Num1, No2, Num2, H1, No, Numx, Numy, S, Nci, whichFloor))
    CaseCXX = CaseCXX + 6
    if float(CaseCXX / LineLen) * 10.0 > Progress:
        Progress = Progress + 1
        print("\r", end = '')
        print('[', end = '')
        for i in range(Progress):
            print('|', end = '')
        for i in range(10 - Progress):
            print(' ', end = '')
        print(']', end = '')
        time.sleep(0.05)
print()
# initial template excel
print('initailize excel...')
NewWb = Workbook()
sheetX = NewWb.active
sheetX.title = '一般柱-X'
NewWb.create_sheet('一般柱-Y')
sheetY = NewWb['一般柱-Y']
for i in range(65,80):
    sheetX.column_dimensions[chr(i)].width = 15.0
    sheetY.column_dimensions[chr(i)].width = 15.0
## SheetX
SheetRange = sheetX['A1':'O1']
iX = 0
for item in SheetRange[0]:
    item.value = InitList_en[iX]
    item.font = Font(name = 'Times New Roman', bold=True, size = 14)
    item.alignment = Alignment(horizontal = 'center')
    iX = iX + 1
SheetRange = sheetX['A2':'O2']
iX = 0
for item in SheetRange[0]:
    item.value = InitList_zh[iX]
    item.font = Font(name = 'Times New Roman', bold=True, size = 14)
    item.alignment = Alignment(horizontal = 'center')
    iX = iX + 1
## SheetY
SheetRange = sheetY['A1':'O1']
iY = 0
for item in SheetRange[0]:
    item.value = InitList_en[iY]
    item.font = Font(name = 'Times New Roman', bold=True, size = 14)
    item.alignment = Alignment(horizontal = 'center')
    iY = iY + 1
SheetRange = sheetY['A2':'O2']
iY = 0
for item in SheetRange[0]:
    item.value = InitList_zh[iY]
    item.font = Font(name = 'Times New Roman', bold=True, size = 14)
    item.alignment = Alignment(horizontal = 'center')
    iY = iY + 1
print('initail excel complete')
## input data to excel
print('writing data to excel...')
CaseCount = 3
Caselen = len(ALLCASE)
Progress = 0
for case in ALLCASE :
    CountStr = str(CaseCount)
    ## X
    sheetX['A' + CountStr] = case.name
    sheetX['B' + CountStr] = case.type
    sheetX['C' + CountStr] = case.HC
    sheetX['D' + CountStr] = case.BC
    sheetX['F' + CountStr] = case.No1
    sheetX['G' + CountStr] = case.Num1  
    sheetX['H' + CountStr] = case.No2  
    sheetX['I' + CountStr] = case.Num2  
    sheetX['J' + CountStr] = case.H1  
    sheetX['K' + CountStr] = case.No  
    sheetX['L' + CountStr] = case.Numx
    sheetX['M' + CountStr] = case.S  
    sheetX['N' + CountStr] = case.Nci  
    sheetX['O' + CountStr] = case.whichFloor
    ## Y
    sheetY['A' + CountStr] = case.name
    sheetY['B' + CountStr] = case.type
    sheetY['C' + CountStr] = case.BC
    sheetY['D' + CountStr] = case.HC
    sheetY['F' + CountStr] = case.No1
    sheetY['G' + CountStr] = case.Num1  
    sheetY['H' + CountStr] = case.No2  
    sheetY['I' + CountStr] = case.Num2  
    sheetY['J' + CountStr] = case.H1  
    sheetY['K' + CountStr] = case.No  
    sheetY['L' + CountStr] = case.Numy
    sheetY['M' + CountStr] = case.S  
    sheetY['N' + CountStr] = case.Nci  
    sheetY['O' + CountStr] = case.whichFloor
    CaseCount = CaseCount + 1
    if float(CaseCount / Caselen) * 10.0 > Progress:
        Progress = Progress + 1
        print("\r", end = '')
        print('[', end = '')
        for i in range(Progress):
            print('|', end = '')
        for i in range(10 - Progress):
            print(' ', end = '')
        print(']', end = '')
        time.sleep(0.05)
print()
print('complete write data to excel\n')
## Final Adjustment
for row in range(2, sheetX.max_row):
    for column in range(sheetX.max_column):
        sheetX.cell(row=row+1,column=column+1).alignment = Alignment(horizontal = 'center')
        sheetX.cell(row=row+1,column=column+1).font = Font(name = 'Times New Roman', size = 14)
for row in range(2, sheetY.max_row):
    for column in range(sheetY.max_column):
        sheetY.cell(row=row+1,column=column+1).alignment = Alignment(horizontal = 'center')
        sheetY.cell(row=row+1,column=column+1).font = Font(name = 'Times New Roman', size = 14)
# handle excption
for item in Num2ExpList:
    print(item + ' has not zero in Num2.')
try:
    NewWb.save('test.xlsx')
    print('\ncomplete!!')
except PermissionError as e:
    print('\nPermission Error. Please close the file and try again.')
os.system('pause')
