import os
# https://openpyxl.readthedocs.io/en/stable/
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.cell import get_column_letter
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment, NamedStyle
from copy import copy
import datetime
import calendar


# https://www.blog.pythonlibrary.org/2021/08/11/styling-excel-cells-with-openpyxl-and-python/#:~:text=Adding%20a%20Border%20OpenPyXL%20gives%20you%20the%20ability,each%20of%20the%20four%20sides%20of%20a%20cell.


#variables
startDate = datetime.datetime(2022, 2, 1)
numberOfMonths = 12
teamMembersLabels = ['Name','Start','End','Monday','Tuesday','Wednessday','Thursday','Friday']
teamMembers = [
                    ['David',datetime.date(2022, 1, 1),None,8,8,8,8,8],
                    ['Maria',datetime.date(2022, 1, 1),None,8,8,8,8,8],
                    ['Gijs',datetime.date(2022, 1, 1),None,8,8,8,8,8],
                    ['Bart',datetime.date(2022, 1, 1),None,8,8,8,8,8],
                    ['Paul',datetime.date(2022, 1, 1),datetime.date(2022, 3, 31),8,0,0,8,0],
                    ['Timothy',datetime.date(2022, 3, 1),None,8,0,0,8,0],
                    ['Michel',datetime.date(2022, 1, 1),None,8,8,8,8,8],
                    ['Jeroen',datetime.date(2022, 1, 1),None,8,8,0,8,8],
                ]

# create a workbook
wb = Workbook()

# Styles
thin = Side(border_style="thin", color="666666")
thin_inactive = Side(border_style="thin", color="AAAAAA")
double = Side(border_style="double", color="666666")

fill_inactive = PatternFill("solid", fgColor="00EEEEEE")
fill_active_header = PatternFill("solid", fgColor="00B7D2FF")
fill_header_label = PatternFill("solid", fgColor="FFFFFF")

font_inactive = Font(color="999999", bold=False)

style_team = NamedStyle(name="team")

style_day = NamedStyle(name="day")
style_day.font = Font(color="000000", bold=True)
style_day.fill = fill_active_header
style_day.alignment = Alignment(horizontal="center", vertical="center")
style_day.border = Border(top=thin, left=thin, right=thin, bottom=thin)


style_day_inactive = copy(style_day)
style_day_inactive.name = "day_inactive"
style_day_inactive.font = font_inactive
style_day_inactive.fill = fill_inactive
style_day_inactive.border = Border(top=thin_inactive, left=thin_inactive, right=thin_inactive, bottom=thin_inactive)

style_weekday = NamedStyle(name="weekday")
style_weekday.alignment = Alignment(horizontal="center", vertical="center")
style_weekday.border = Border(top=thin, left=thin, right=thin, bottom=thin)
style_weekday.fill = fill_active_header

style_weekday_inactive = copy(style_day)
style_weekday_inactive.name = "weekday_inactive"
style_weekday_inactive.font = font_inactive
style_weekday_inactive.fill = fill_inactive
style_weekday_inactive.border = Border(top=thin_inactive, left=thin_inactive, right=thin_inactive, bottom=thin_inactive)


style_month = NamedStyle(name="month")
style_month.font = Font(color="000000", bold=True)
style_month.border = Border(top=thin, left=thin, right=thin, bottom=thin)
style_month.alignment = Alignment(horizontal="center", vertical="center")
style_month.fill = fill_active_header

style_month_inactive = copy(style_month)
style_month_inactive.name = 'month_inactive'
style_month_inactive.font = font_inactive
style_month_inactive.fill = fill_inactive
style_month_inactive.border = Border(top=thin_inactive, left=thin_inactive, right=thin_inactive, bottom=thin_inactive)

style_team_header = NamedStyle(name="style_team_header")
style_team_header.font = Font(color="000000", bold=True)
# style_team_header.border = Border(top=thin, left=thin, right=thin, bottom=thin)
style_team_header.alignment = Alignment(horizontal="center", vertical="center")
style_team_header.fill = fill_header_label

style_team = NamedStyle(name="style_team")
style_team.font = Font(color="000000", bold=False)
style_team.border = Border(top=thin, left=thin, bottom=thin)
style_team.alignment = Alignment(horizontal="left", vertical="center")
style_team.fill = fill_active_header

style_team_inactive = copy(style_team)
style_team_inactive.name = "style_team_inactive"
style_team_inactive.font = font_inactive
style_team_inactive.border = Border(top=thin_inactive, left=thin_inactive, bottom=thin_inactive)
style_team_inactive.alignment = Alignment(horizontal="left", vertical="center")
style_team_inactive.fill = fill_inactive


def addMonths(d, x):
    newday = d.day
    newmonth = (((d.month - 1) + x) % 12) + 1
    newyear  = d.year + (((d.month - 1) + x) // 12)
    if newday > calendar.mdays[newmonth]:
        newday = calendar.mdays[newmonth]
        if newyear % 4 == 0 and newmonth == 2:
            newday += 1
    return datetime.date(newyear, newmonth, newday)





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



# creating the month sheets
for monthNumber in range(0,numberOfMonths):
    # mr = calendar.monthrange(2022,13)
    date = addMonths(startDate,monthNumber)
    # yearNumber = startDate.Add()

    ws = wb.create_sheet(date.strftime('%Y-%m'))
    startRowNumber = 4
    ws['A{0}'.format(startRowNumber)]='Team'
    ws['A{0}'.format(startRowNumber)].style = style_team_header
    counter = 0
    for row in teamMembers:
        counter += 1
        ws['A{0}'.format(startRowNumber+counter)]=row[0]
        if (row[1] is None or row[1] <= date) and (row[2] is None or row[2] > date):
            ws['A{0}'.format(startRowNumber+counter)].style = style_team
        else:
            ws['A{0}'.format(startRowNumber+counter)].style = style_team_inactive
        
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
        ws['{0}{1}'.format(col,2)]=day.strftime('%b') # Month as locale’s abbreviated name.  (https://strftime.org/)
        ws['{0}{1}'.format(col,4)]=day.strftime('%a') # Weekday as locale’s abbreviated name. (https://strftime.org/)

        ws['{0}{1}'.format(col,3)]=day.day
        if ( dayCounter> endHeaderDays and dayCounter < startTailDates  ):
            ws['{0}{1}'.format(col,3)].style = style_day
            ws['{0}{1}'.format(col,2)].style = style_month
            ws['{0}{1}'.format(col,4)].style = style_weekday
        else:
            ws['{0}{1}'.format(col,3)].style = style_day_inactive
            ws['{0}{1}'.format(col,2)].style = style_month_inactive
            ws['{0}{1}'.format(col,4)].style = style_weekday_inactive

        ws.column_dimensions[col].width = 5

    # merge month header 
    if (endHeaderDays > startHeaderDays):
        ws.merge_cells('{0}{1}:{2}{3}'.format(get_column_letter(2),2,get_column_letter(1 + date.weekday()),2))
    
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
filename = "output\plan.xlsx"
wb.save(filename)

os.system("{}".format(filename))