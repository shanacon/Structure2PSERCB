from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.styles import Alignment


InitList_en = ['name', 'type', 'Bc/Dc', 'Hc', 'Los(%)', 'No1', 'Num1', 'No2', 'Num2', 'h1', 'No', 'Num', 'S', 'Nci', 'whichFloor']
InitList_zh= ['斷面名稱', '型式', '寬度/直徑', '深度', '主筋鋼筋比', '主筋號數(1)', '主筋根數(1)', '主筋號數(2)', '主筋根數(2)', 
              '一樓柱淨高', '橫向箍、繫筋號數', '橫向箍、繫筋根數', '箍筋間距', '柱根數', '所在樓層']
## initial template excel
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

## Final Adjustment
for row in range(2, sheetX.max_row):
    for column in range(sheetX.max_column):
        sheetX.cell(row=row+1,column=column+1).alignment = Alignment(horizontal = 'center')
        sheetX.cell(row=row+1,column=column+1).font = Font(name = 'Times New Roman')
NewWb.save('test.xlsx')