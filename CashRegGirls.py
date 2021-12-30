import openpyxl
from openpyxl.styles import NamedStyle, Border, Side, Font, GradientFill, Alignment, PatternFill, Color
from openpyxl.writer.excel import save_workbook
# from openpyxl.utils import FORMULAE
from openpyxl.chart import BarChart, Reference
import pandas as pd
import easygui

# Reminder of the correct file names
ip_list = sorted(["ВИА", "ШИИ", "РИЭ"])
files = easygui.msgbox('Работа программы потребует наличие двух файлов для каждого ИП:\n\
ФИО_чеки\nФИО_часы')

# Import and process data from files 'Checks'
# ВИА files
df = pd.read_excel('ВИА_чеки.xls', header = None)
df.to_excel('ВИА_чеки.xlsx', index = False, header = False)
wbb1 = openpyxl.load_workbook('ВИА_чеки.xlsx')
sh_VIAchex = wbb1.worksheets[0]
chexNumVIA = sh_VIAchex.max_row-3

# Time Intervals Stats VIA

timeIntervalsTuple = ('08:30','09:00', '09:30', '10:00', '10:30', '11:00', '11:30', '12:00', '12:30', '13:00', '13:30',
                      '14:00', '14:30', '15:00', '15:30', '16:00', '16:30', '17:00', '17:30', '18:00', '18:30',
                      '19:00', '19:30', '20:00')

# A list of time intervals as keys
timeIntervalsVIA = dict.fromkeys(timeIntervalsTuple, 0)

for i in range(4, sh_VIAchex.max_row+1):
    t = str(sh_VIAchex.cell(row=i, column=4).value)
    t1 = t[:4]
    t2 = t[:3] + '3'
    if t1 < t2:
        timeIntervalsVIA[t[:3]+'00'] += 1
    else:
        timeIntervalsVIA[t[:3]+'30'] += 1

# ++++++++++++++++++++++++++++============================================================

girlsVIA = []  # Number and list of girls
for i in range(4, sh_VIAchex.max_row+1):
    g = sh_VIAchex.cell(row = i, column = 6).value
    if g != None and g not in girlsVIA:
        girlsVIA.append(g)
    elif g is None:
        x = 'Аноним'
        if x not in girlsVIA:
            girlsVIA.append(x)

girlsNumVIA = len(girlsVIA)


orderNumVIA = 0  # Order qty
for i in range(4, sh_VIAchex.max_row+1):
    order = sh_VIAchex.cell(row = i, column = 7).value
    if order == 1:
        orderNumVIA += 1

# VIA Girls' personal statistics
girlsVIAchex = dict.fromkeys(girlsVIA, 0)
girlsVIAorders = dict.fromkeys(girlsVIA, 0)
girlsVIAsums = dict.fromkeys(girlsVIA, int(0))

for i in range(4, sh_VIAchex.max_row+1):
    name = sh_VIAchex.cell(row=i, column=6).value
    if name is None:
        name = 'Аноним'
    order = sh_VIAchex.cell(row=i, column=7).value
    checkSum = sh_VIAchex.cell(row=i, column=8).value
    girlsVIAchex[name] += 1
    girlsVIAorders[name] += order
    girlsVIAsums[name] += int(checkSum)

valSort_girlsVIAchex = sorted(girlsVIAchex.values(), reverse=True) # girlsVIAchex sorted by values
girlsVIAchexSorted = {}
for i in valSort_girlsVIAchex:
    for k in girlsVIAchex.keys():
        if girlsVIAchex[k] == i:
            girlsVIAchexSorted[k] = girlsVIAchex[k]

# Identifying period (month, year)
dateCell=str(sh_VIAchex['C4'].value)
monthNum = dateCell[3:5]
year = dateCell[6:8]
month_list={'01':'Январь', '02':'Февраль', '03':'Март', '04':'Апрель', '05':'Май', '06':'Июнь',
            '07':'Июль', '08':'Август', '09':'Сентябрь', '10':'Октябрь', '11':'Ноябрь', '12':'Декабрь'}
monthName=month_list[monthNum]
period=str(monthName+' 20'+year)

wbb1.close()

# РИЭ files
df = pd.read_excel('РИЭ_чеки.xls', header = None)
df.to_excel('РИЭ_чеки.xlsx', index = False, header = False)
wbb2 = openpyxl.load_workbook('РИЭ_чеки.xlsx')
sh_RIEchex = wbb2.worksheets[0]
chexNumRIE = sh_RIEchex.max_row-3


# Time Intervals Stats RIE
# A list of time intervals as keys
timeIntervalsRIE = dict.fromkeys(timeIntervalsTuple, 0)

for i in range(4, sh_RIEchex.max_row+1):
    t = str(sh_RIEchex.cell(row=i, column=4).value)
    t1 = t[:4]
    t2 = t[:3] + '3'
    if t1 < t2:
        timeIntervalsRIE[t[:3]+'00'] += 1
    else:
        timeIntervalsRIE[t[:3]+'30'] += 1

# RIE Name List
girlsRIE = []
for i in range(4, sh_RIEchex.max_row+1):
    g = sh_RIEchex.cell(row = i, column = 6).value
    if g == None:
        g = 'Аноним'
        if g not in girlsRIE:
            girlsVIA.append(g)
    elif g not in girlsRIE:
        girlsRIE.append(g)
girlsNumRIE = len(girlsRIE)

orderNumRIE = 0  # Order qty
for i in range(4, sh_RIEchex.max_row+1):
    order = sh_RIEchex.cell(row = i, column = 7).value
    if order == 1:
        orderNumRIE += 1

# RIE Girls' personal statistics
girlsRIEchex = dict.fromkeys(girlsRIE, 0)
girlsRIEorders = dict.fromkeys(girlsRIE, 0)
girlsRIEsums = dict.fromkeys(girlsRIE, int(0))

for i in range(4, sh_RIEchex.max_row+1):
    name = sh_RIEchex.cell(row=i, column=6).value
    if name is None:
        name = 'Аноним'
    order = sh_RIEchex.cell(row=i, column=7).value
    checkSum = sh_RIEchex.cell(row=i, column=8).value
    girlsRIEchex[name] += 1
    girlsRIEorders[name] += order
    girlsRIEsums[name] += int(checkSum)

valSort_girlsRIEchex = sorted(girlsRIEchex.values(), reverse=True) # girlsRIEchex sorted by values
girlsRIEchexSorted = {}
for i in valSort_girlsRIEchex:
    for k in girlsRIEchex.keys():
        if girlsRIEchex[k] == i:
            girlsRIEchexSorted[k] = girlsRIEchex[k]


wbb2.close()

# ШИИ files
df = pd.read_excel('ШИИ_чеки.xls', header = None)
df.to_excel('ШИИ_чеки.xlsx', index = False, header = False)
wbb3 = openpyxl.load_workbook('ШИИ_чеки.xlsx')
sh_SHIIchex = wbb3.worksheets[0]
chexNumSHII = sh_SHIIchex.max_row-3

# Time Intervals Stats SHII
# A list of time intervals as keys
timeIntervalsSHII = dict.fromkeys(timeIntervalsTuple, 0)

for i in range(4, sh_SHIIchex.max_row+1):
    t = str(sh_SHIIchex.cell(row=i, column=4).value)
    t1 = t[:4]
    t2 = t[:3] + '3'
    if t1 < t2:
        timeIntervalsSHII[t[:3]+'00'] += 1
    else:
        timeIntervalsSHII[t[:3]+'30'] += 1

# SHII Name List
girlsSHII = []
for i in range(4, sh_SHIIchex.max_row+1):
    g = sh_SHIIchex.cell(row = i, column = 6).value
    if g is None:
        x = 'Аноним'
        if x not in girlsSHII:
            girlsSHII.append(x)
    elif g not in girlsSHII:
        girlsSHII.append(g)
girlsNumSHII = len(girlsSHII)

orderNumSHII = 0  # Order qty
for i in range(4, sh_SHIIchex.max_row+1):
    order = sh_SHIIchex.cell(row = i, column = 7).value
    if order == 1:
        orderNumSHII += 1

# SHII Girls' personal statistics
girlsSHIIchex = dict.fromkeys(girlsSHII, 0)
girlsSHIIorders = dict.fromkeys(girlsSHII, 0)
girlsSHIIsums = dict.fromkeys(girlsSHII, int(0))

for i in range(4, sh_SHIIchex.max_row+1):
    name = sh_SHIIchex.cell(row=i, column=6).value
    if name is None:
        name = 'Аноним'
    order = sh_SHIIchex.cell(row=i, column=7).value
    checkSum = sh_SHIIchex.cell(row=i, column=8).value
    girlsSHIIchex[name] += 1
    girlsSHIIorders[name] += order
    girlsSHIIsums[name] += int(checkSum)

valSort_girlsSHIIchex = sorted(girlsSHIIchex.values(), reverse=True) # girlsSHIIchex sorted by values
girlsSHIIchexSorted = {}
for i in valSort_girlsSHIIchex:
    for k in girlsSHIIchex.keys():
        if girlsSHIIchex[k] == i:
            girlsSHIIchexSorted[k] = girlsSHIIchex[k]
            break

wbb3.close()

# Import and process data from files 'Hours'

df = pd.read_excel('ШИИ_чеки.xls', header = None)
df.to_excel('ШИИ_чеки.xlsx', index = False, header = False)
wbb3 = openpyxl.load_workbook('ШИИ_чеки.xlsx')

sh_SHIIchex = wbb3.worksheets[0]


# Import and process data from files 'Orders'
# ВИА files
"""df = pd.read_excel('ВИА_заказы.xls', header = None)
df.to_excel('ВИА_заказы.xlsx', index = False, header = False)
wbb1 = openpyxl.load_workbook('ВИА_заказы.xlsx')
sh_VIAorder = wbb1.worksheets[0]

cntName = 0
name = str('Кассир')
for i in range(1, sh_VIAorder.max_row+1):
    names = str(sh_VIAorder.cell(row = i, column = 5).value)
    if name in names:
        cntName += 1
orderNumVIA = cntName

wbb1.close()  

# РИЭ files
df = pd.read_excel('РИЭ_заказы.xls', header = None)
df.to_excel('РИЭ_заказы.xlsx', index = False, header = False)
wbb2 = openpyxl.load_workbook('РИЭ_заказы.xlsx')
sh_RIEorder = wbb2.worksheets[0]

cntName = 0
name = str('Кассир')
for i in range(1, sh_RIEorder.max_row+1):
    names = str(sh_RIEorder.cell(row = i, column = 5).value)
    if name in names:
        cntName += 1
orderNumRIE = cntName

wbb2.close()  

# ШИИ files
df = pd.read_excel('ШИИ_заказы.xls', header = None)
df.to_excel('ШИИ_заказы.xlsx', index = False, header = False)
wbb3 = openpyxl.load_workbook('ШИИ_заказы.xlsx')
sh_SHIIorder = wbb3.worksheets[0]

cntName = 0
name = str('Кассир')
for i in range(1, sh_SHIIorder.max_row+1):
    names = str(sh_SHIIorder.cell(row = i, column = 5).value)
    if name in names:
        cntName += 1
orderNumSHII = cntName   

wbb3.close()  """


# wbb4 = xlsxwriter.Workbook('Нагрузка кассы.xlsx')
# wsh4 = wbb4.add_worksheet(period)
# wsh4.set_margins(left=0.6, right=0.2, top=0.2, bottom=0.2)
# wbb4.close()

workloadFile = 'Нагрузка кассы.xlsx'
try:
    wbb4 = openpyxl.load_workbook(workloadFile)
except:
    openpyxl.Workbook()
wsh4 = wbb4.create_sheet(period)

# Cell style
boldFont = Font(bold=True, size=10)
plainFont = Font(bold=False, size=10)
alignCenter = Alignment(horizontal='center', vertical='center')
alignLeft = Alignment(horizontal='left', vertical='center')
borderLines = Side(border_style='thin', color='000000')
squareBorder = Border(top=borderLines,
                      bottom=borderLines,
                      right=borderLines,
                      left=borderLines)
cellFillYellow = PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid')
cellFillGreenish = PatternFill(start_color='CCFF99', end_color='CCFF99', fill_type='solid')


# Drawing the table
for i in range(3, 7):
    for j in range(1, 10):
        wsh4.cell(row=i, column=j).border = squareBorder
        wsh4.cell(row=i, column=j).alignment = alignCenter
        if i == 3:
            wsh4.cell(row=i, column=j).font = boldFont
            wsh4.cell(row=i, column=j).fill = cellFillYellow
        if j == 1 and i != 3:
            wsh4.cell(row=i, column=j).fill = cellFillGreenish

wsh4.column_dimensions['A'].width = 25
cellList = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
for j in range(2, 7):
    wsh4.row_dimensions[j].height = 25
for i in cellList:
    wsh4.column_dimensions[i].width = 15

wsh4['E2'] = 'НАГРУЗКА НА КАССЫ'
wsh4['E2'].font = boldFont
wsh4['A3'] = period
wsh4['A4'].alignment = alignLeft
wsh4['A5'].alignment = alignLeft
wsh4['A6'].alignment = alignLeft
wsh4['A4'] = '  ИП Вербовская И.А.'
wsh4['A5'] = '  ИП Рейн И.Э.'
wsh4['A6'] = '  ИП Ширяев И.И.'
headerList = ['Кол-во сотр.', 'Кол-во чек.', 'Заказов', 'Всего часов',
              'Чек / сотр.', 'Заказ / сотр.', 'Час / сотр.', 'Чек / час']
col = 2
for i in headerList:
    wsh4.cell(row=3, column=col).value = i
    col += 1

for i in range(4, 7):
    for j in range(1, 10):
        wsh4.cell(row=i, column=j).font = plainFont


# Inserting data
# wsh4.cell(row=4, column=6).value="{:,}".format(wsh4['F4']).replace(',', ' ')
wsh4['B4'] = girlsNumVIA
wsh4['B5'] = girlsNumRIE
wsh4['B6'] = girlsNumSHII
wsh4['C4'] = chexNumVIA
wsh4['C5'] = chexNumRIE
wsh4['C6'] = chexNumSHII
wsh4['D4'] = orderNumVIA
wsh4['D5'] = orderNumRIE
wsh4['D6'] = orderNumSHII

wsh4['F4'] = round(chexNumVIA /  girlsNumVIA, 0)
wsh4['F5'] = round(chexNumRIE /  girlsNumRIE, 0)
wsh4['F6'] = round(chexNumSHII /  girlsNumSHII, 0)
wsh4['G4'] = round(orderNumVIA /  girlsNumVIA, 1)
wsh4['G5'] = round(orderNumRIE /  girlsNumRIE, 1)
wsh4['G6'] = round(orderNumSHII /  girlsNumSHII, 1)
wsh4['H4'] = '=E4 / B4'
wsh4['H5'] = '=E5 / B5'
wsh4['H6'] = '=E6 / B6'
wsh4['I4'] = '=C4 / E4'
wsh4['I5'] = '=C5 / E5'
wsh4['I6'] = '=C6 / E6'

# wsh4['F4'].value="{:,}".format().replace(',', ' ')



# Girls' personal stats part

# Function for returning dict. key by value
def getKeyVIA(val):
    for key, value in girlsVIAchexSorted.items():
        if val == value:
            return key

row = 10
for i in girlsVIAchexSorted.values():  # Inserting list sorted by check qty
    wsh4.cell(row=row, column=3).value = i
    wsh4.cell(row=row, column=1).value = getKeyVIA(i)
    wsh4.cell(row=row, column=4).value = girlsVIAorders.get(getKeyVIA(i))
    row += 1

# Function for returning dict. key by value
def getKeyRIE(val):
    for key, value in girlsRIEchexSorted.items():
        if val == value:
            return key

row +=2
for i in girlsRIEchexSorted.values():  # Inserting list sorted by check qty
    wsh4.cell(row=row, column=3).value = i
    wsh4.cell(row=row, column=1).value = getKeyRIE(i)
    wsh4.cell(row=row, column=4).value = girlsRIEorders.get(getKeyRIE(i))
    row += 1


# Function for returning dict. key by value
def getKeySHII(val):
    for key, value in girlsSHIIchexSorted.items():
        if val == value:
            return key

row +=2
for i in girlsSHIIchexSorted.values():  # Inserting list sorted by check qty
    wsh4.cell(row=row, column=3).value = i
    wsh4.cell(row=row, column=1).value = getKeySHII(i)
    wsh4.cell(row=row, column=4).value = girlsSHIIorders.get(getKeySHII(i))
    row += 1



# VIA Chart Data
row += 2
row2 = row
wsh4.cell(row=row, column=8).value = 'БД1'
wsh4.cell(row=row, column=9).value = 'БД3'
wsh4.cell(row=row, column=10).value = 'БД4'
row += 1

for i in timeIntervalsVIA.keys():
    j = timeIntervalsVIA[i]
    k = timeIntervalsRIE[i]
    l = timeIntervalsSHII[i]
    wsh4.cell(row=row, column=7).value = i
    wsh4.cell(row=row, column=8).value = j
    wsh4.cell(row=row, column=9).value = k
    wsh4.cell(row=row, column=10).value = l
    row += 1
row -= 1



# Bar Chart
timeIntervalsChart = BarChart()
chartCats = Reference(worksheet=wsh4,
                      min_row=row2+1, max_row=row,
                      min_col=7, max_col=7)
chartData = Reference(worksheet=wsh4,
                      min_row=row2, max_row=row,
                      min_col=8, max_col=10)
timeIntervalsChart.add_data(chartData, titles_from_data=True)
timeIntervalsChart.x_axis.title = "Временные интервалы"
timeIntervalsChart.y_axis.title = "Кол-во чеков по ИП"
timeIntervalsChart.width = 23
wsh4.add_chart(timeIntervalsChart, "A57")
timeIntervalsChart.set_categories(chartCats)


wsh4.print_area = 'A1:I77'
wbb4.active = wsh4
wbb4.save(str(workloadFile))
wbb4.close()

print(period)
