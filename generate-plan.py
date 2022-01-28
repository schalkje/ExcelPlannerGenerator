# https://openpyxl.readthedocs.io/en/stable/
from asyncio.windows_events import NULL
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.cell import get_column_letter
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
import datetime
import calendar

#variables
startDate = datetime.datetime(2022, 2, 1)
numberOfMonths = 12
teamMembersLabels = ['Name','Start','End','Monday','Tuesday','Wednessday','Thursday','Friday']
teamMembers = [
                    ['David',datetime.date(2022, 1, 1),NULL,8,8,8,8,8],
                    ['Maria',datetime.date(2022, 1, 1),NULL,8,8,8,8,8],
                    ['Gijs',datetime.date(2022, 1, 1),NULL,8,8,8,8,8],
                    ['Bart',datetime.date(2022, 1, 1),NULL,8,8,8,8,8],
                    ['Paul',datetime.date(2022, 1, 1),datetime.date(2022, 3, 31),8,0,0,8,0],
                    ['Timothy',datetime.date(2022, 3, 1),NULL,8,0,0,8,0],
                    ['Michel',datetime.date(2022, 1, 1),NULL,8,8,8,8,8],
                    ['Jeroen',datetime.date(2022, 1, 1),NULL,8,8,0,8,8],
                ]


def addMonths(d, x):
    newday = d.day
    newmonth = (((d.month - 1) + x) % 12) + 1
    newyear  = d.year + (((d.month - 1) + x) // 12)
    if newday > calendar.mdays[newmonth]:
        newday = calendar.mdays[newmonth]
        if newyear % 4 == 0 and newmonth == 2:
            newday += 1
    return datetime.date(newyear, newmonth, newday)



# create a workbook
wb = Workbook()

generatorVariables = [
    ['startDate',startDate],
    ['numberOfMonths',numberOfMonths],
]


# grab the active worksheet and rename
ws = wb.active
ws.title = "Overview"

ws.append(['Variable', 'Value'])
ws.column_dimensions['A'].width = 22
ws.column_dimensions['B'].width = 30

for row in generatorVariables:
    ws.append(row)

tab = Table(displayName="GeneratorVariables", ref="A1:B3")
# Add a default style with striped rows and banded columns
style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                       showLastColumn=False, showRowStripes=True, showColumnStripes=True)
tab.tableStyleInfo = style

ws.add_table(tab)

# empty line
ws.append([''])

startRowNumber = 5
ws.append(teamMembersLabels)
ws.column_dimensions['A'].width = 22
ws.column_dimensions['B'].width = 30
ws.column_dimensions['C'].width = 30
ws.column_dimensions['D'].width = 10
ws.column_dimensions['E'].width = 10
ws.column_dimensions['F'].width = 10
ws.column_dimensions['G'].width = 10
ws.column_dimensions['H'].width = 10

for row in teamMembers:
    ws.append(row)

tab = Table(displayName="TeamMembers", ref="A{0}:H{1}".format(startRowNumber,startRowNumber+len(teamMembers)))
# Add a default style with striped rows and banded columns
style = TableStyleInfo(name="TableStyleMedium11", showFirstColumn=True,
                       showLastColumn=False, showRowStripes=True, showColumnStripes=False)
tab.tableStyleInfo = style

ws.add_table(tab)




for monthNumber in range(0,numberOfMonths):
    # mr = calendar.monthrange(2022,13)
    date = addMonths(startDate,monthNumber)
    # yearNumber = startDate.Add()

    ws = wb.create_sheet(date.strftime('%Y-%m'))
    startRowNumber = 5
    ws['A{0}'.format(startRowNumber)]='Name'
    counter = 0
    for row in teamMembers:
        counter += 1
        if (row[1] is NULL or row[1] <= date) and (row[2] is NULL or row[2] > date):
                ws['A{0}'.format(startRowNumber+counter)]=row[0]
    cal = calendar.Calendar(0)
    days = cal.itermonthdates(date.year, date.month)
    startHeaderDays = 2
    endHeaderDays = 1 + date.weekday()
    startTailDates = 2 + date.weekday()+calendar.monthrange(date.year,date.month)[1]
    endTailDates = (int( (startTailDates - 1) / 7 ) + 1) * 7 + 1
    startColumn = 1
    dayCounter = startHeaderDays -1
    for day in days:
        startColumn += 1
        dayCounter += 1
        col = get_column_letter(startColumn)
        # https://www.blog.pythonlibrary.org/2021/08/11/styling-excel-cells-with-openpyxl-and-python/#:~:text=Adding%20a%20Border%20OpenPyXL%20gives%20you%20the%20ability,each%20of%20the%20four%20sides%20of%20a%20cell.
        ws['{0}{1}'.format(col,2)]=day.strftime('%b')
        ws['{0}{1}'.format(col,2)].font = Font(color="FF0000")
        ws['{0}{1}'.format(col,3)]=day.day
        # ws['{0}{1}'.format(col,3)].border = Border(top='double', left='thin', right='thin', bottom='double')
        if ( dayCounter> endHeaderDays ):
            ws['{0}{1}'.format(col,3)].fill = PatternFill("solid", fgColor="DDDD00")
        else:
            ws['{0}{1}'.format(col,3)].fill = PatternFill("solid", fgColor="DDDDDD")
        ws['{0}{1}'.format(col,3)].font = Font(bold=True)
        ws['{0}{1}'.format(col,3)].alignment = Alignment(horizontal="center", vertical="center")
        ws['{0}{1}'.format(col,4)]=day.strftime('%a')
        ws['{0}{1}'.format(col,4)].alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[col].width = 5
    if (endHeaderDays > startHeaderDays):
        ws.merge_cells('{0}{1}:{2}{3}'.format(get_column_letter(2),2,get_column_letter(1 + date.weekday()),2))
    # col = 2 + date.weekday()+calendar.monthrange(date.year,date.month)[1]
    ws.merge_cells('{0}{1}:{2}{3}'.format(get_column_letter(2 + date.weekday()),2,get_column_letter(1 + date.weekday()+calendar.monthrange(date.year,date.month)[1]),2))
    
    if (startTailDates < endTailDates):
        ws.merge_cells('{0}{1}:{2}{3}'.format(get_column_letter(startTailDates),2,get_column_letter(endTailDates),2))
    


# # Data can be assigned directly to cells
# ws['A1'] = 42

# # Rows can also be appended
# ws.append([1, 2, 3])

# # Python types will automatically be converted
# ws['A2'] = datetime.datetime.now()

# Save the file
wb.save("output\plan.xlsx")