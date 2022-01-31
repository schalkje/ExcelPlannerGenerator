import os
# https://openpyxl.readthedocs.io/en/stable/
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.cell import get_column_letter
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment, NamedStyle
from openpyxl.worksheet.properties import WorksheetProperties, PageSetupProperties
from copy import copy
import datetime
import calendar


# https://www.blog.pythonlibrary.org/2021/08/11/styling-excel-cells-with-openpyxl-and-python/#:~:text=Adding%20a%20Border%20OpenPyXL%20gives%20you%20the%20ability,each%20of%20the%20four%20sides%20of%20a%20cell.


#variables
startDate = datetime.date(2022, 2, 1)
numberOfMonths = 12
first_flip_day = datetime.date(2022, 2, 8)
sprint_size = 14 # (2 weeks)
num_sprints = 20
flip_days = [first_flip_day + datetime.timedelta(days=(x*sprint_size)) for x in range(num_sprints)]
teamMembersLabels = ['Name','Start','End','Monday','Tuesday','Wednessday','Thursday','Friday']
teamMembers = [
                    ['David',datetime.date(2022, 1, 1),None,8,8,8,8,8],
                    ['Maria',datetime.date(2022, 1, 1),None,8,8,8,8,8],
                    ['Gijs',datetime.date(2022, 1, 1),None,8,8,8,8,8],
                    ['Bart',datetime.date(2022, 1, 1),None,8,8,8,8,8],
                    ['Paul',datetime.date(2022, 1, 1),datetime.date(2022, 3, 31),8,0,0,8,0],
                    ['Timothy',datetime.date(2022, 3, 1),None,8,8,8,8,0],
                    ['Michel',datetime.date(2022, 1, 1),None,8,8,8,8,8],
                    ['Jeroen',datetime.date(2022, 1, 1),None,8,8,0,8,8],
                ]

# create a workbook
wb = Workbook()

# Styles
# thin = Side(border_style="thin", color="666666")
# thin_inactive = Side(border_style="thin", color="AAAAAA")
thin = Side(border_style="medium", color="FFFFFF")
thin_inactive = Side(border_style="medium", color="FFFFFF")
double = Side(border_style="double", color="666666")
side_team = Side(border_style="thick", color="FFFFFF")
side_team_vertical = Side(border_style="medium", color="FFFFFF")

fill_inactive = PatternFill("solid", fgColor="00EEEEEE")
fill_team_inactive = PatternFill("solid", fgColor="00FFFFFF")
fill_active_header = PatternFill("solid", fgColor="00B7D2FF")
fill_header_label = PatternFill("solid", fgColor="FFFFFF")
fill_weekend = PatternFill("solid", fgColor="00FFFFFF")
fill_workday_odd = PatternFill("solid", fgColor="00DCE6F1")
fill_workday_even = PatternFill("solid", fgColor="00B8CCE4")
fill_footer = PatternFill("solid", fgColor="00DBF1DE")
fill_flip_day= PatternFill("solid", fgColor="00DFE18F")

font_inactive = Font(color="999999", bold=False)

style_team = NamedStyle(name="team")

style_day = NamedStyle(name="day")
style_day.font = Font(color="000000", bold=True)
style_day.fill = fill_active_header
style_day.alignment = Alignment(horizontal="center", vertical="center")
style_day.border = Border(top=thin, left=thin, right=thin, bottom=thin)


style_flip_day = NamedStyle(name="flip_day")
style_flip_day.font = Font(color="000000", bold=True)
style_flip_day.fill = fill_flip_day
style_flip_day.alignment = Alignment(horizontal="center", vertical="center")
style_flip_day.border = Border(top=thin, left=thin, right=thin, bottom=thin)

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

style_weekend = NamedStyle(name="weekend")
style_weekend.alignment = Alignment(horizontal="center", vertical="center")
style_weekend.border = Border(bottom=side_team)
style_weekend.font = font_inactive
style_weekend.fill = fill_weekend

style_weekend_inactive = copy(style_weekend)
style_weekend_inactive.name = "weekend_inactive"
style_weekend_inactive.font = font_inactive
style_weekend_inactive.fill = fill_weekend


style_workday_odd = NamedStyle(name="workday_odd")
style_workday_odd.alignment = Alignment(horizontal="center", vertical="center")
style_workday_odd.border = Border(bottom=side_team, left=side_team_vertical, right=side_team_vertical)
style_workday_odd.fill = fill_workday_odd

style_team_odd = copy(style_workday_odd)
style_team_odd.name = "style_team_odd"
style_team_odd.alignment = Alignment(horizontal="left", vertical="center")
style_team_odd.font = Font(bold=True)

style_workday_even = NamedStyle(name="workday_even")
style_workday_even.alignment = Alignment(horizontal="center", vertical="center")
style_workday_even.border = Border(bottom=side_team, left=side_team_vertical, right=side_team_vertical)
style_workday_even.fill = fill_workday_even

style_team_even = copy(style_workday_even)
style_team_even.name = "style_team_even"
style_team_even.alignment = Alignment(horizontal="left", vertical="center")
style_team_even.font = Font(bold=True)

style_workday_inactive = copy(style_day)
style_workday_inactive.name = "workday_inactive"
style_workday_inactive.font = font_inactive
style_workday_inactive.fill = fill_inactive
style_workday_inactive.border = Border(bottom=side_team, left=side_team_vertical, right=side_team_vertical)


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
style_team_header.alignment = Alignment(horizontal="center", vertical="bottom")
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
style_team_inactive.fill = fill_team_inactive

style_team_workday_inactive = copy(style_team)
style_team_workday_inactive.name = "style_team_workday_inactive"
style_team_workday_inactive.font = font_inactive
style_team_workday_inactive.border = Border(top=thin_inactive, left=thin_inactive, bottom=thin_inactive)
style_team_workday_inactive.alignment = Alignment(horizontal="center", vertical="center")
style_team_workday_inactive.fill = fill_team_inactive


style_day_sum = NamedStyle(name="style_day_sum")
style_day_sum.font = Font(color="000000", bold=False)
style_day_sum.border = Border(top=thin, left=thin, bottom=thin)
style_day_sum.alignment = Alignment(horizontal="center", vertical="center")
style_day_sum.fill = fill_footer

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
ws_overview = wb.active
ws_overview.title = "Overview"
props = ws_overview.sheet_properties
props.tabColor = "1072BA"

ws_overview.append(['Variable', 'Value'])
ws_overview.column_dimensions['A'].width = 22
ws_overview.column_dimensions['B'].width = 30

for row in generatorVariables:
    ws_overview.append(row)

tab = Table(displayName="GeneratorVariables", ref="A1:B3")
# Add a default style with striped rows_overview and banded columns
style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                       showLastColumn=False, showRowStripes=True, showColumnStripes=True)
tab.tableStyleInfo = style

ws_overview.add_table(tab)

# empty line
ws_overview.append([''])

startRowNumber = 5
ws_overview.append(teamMembersLabels)
ws_overview.column_dimensions['A'].width = 22
ws_overview.column_dimensions['B'].width = 30
ws_overview.column_dimensions['C'].width = 30
ws_overview.column_dimensions['D'].width = 10
ws_overview.column_dimensions['E'].width = 10
ws_overview.column_dimensions['F'].width = 10
ws_overview.column_dimensions['G'].width = 10
ws_overview.column_dimensions['H'].width = 10

for row in teamMembers:
    ws_overview.append(row)

tab = Table(displayName="TeamMembers", ref="A{0}:H{1}".format(startRowNumber,startRowNumber+len(teamMembers)))
# Add a default style with striped rows_overview and banded columns
style = TableStyleInfo(name="TableStyleMedium11", showFirstColumn=True,
                       showLastColumn=False, showRowStripes=True, showColumnStripes=False)
tab.tableStyleInfo = style

ws_overview.add_table(tab)



# creating the month sheets
for monthNumber in range(0,numberOfMonths):
    # mr = calendar.monthrange(2022,13)
    date = addMonths(startDate,monthNumber)
    # yearNumber = startDate.Add()

    ws = wb.create_sheet(title=date.strftime('%Y-%m'))
    # props = ws.sheet_properties
    ws.sheet_view.showGridLines = False
    # ws.sheet_view.showRowColHeaders = False


        
    cal = calendar.Calendar(0)
    days = cal.itermonthdates(date.year, date.month)

    offset_rows = 1
    offset_cols = 3
    team_offset_cols = 2
    team_offset_rows = 5    

    startHeaderDays = 1
    endHeaderDays = date.weekday()
    endMonthDays = date.weekday()+calendar.monthrange(date.year,date.month)[1]
    endTailDates = (int( (endMonthDays) / 7 ) + 1) * 7 
    column_nr = offset_cols
    dayCounter = 0
    for day in days:
        column_nr += 1
        dayCounter += 1
        col = get_column_letter(column_nr)
        ws['{0}{1}'.format(col,2)]=day.strftime('%b') # Month as locale’s abbreviated name.  (https://strftime.org/)
        ws['{0}{1}'.format(col,4)]=day.strftime('%a') # Weekday as locale’s abbreviated name. (https://strftime.org/)

        ws['{0}{1}'.format(col,3)]=day.day
        if ( dayCounter > endHeaderDays and dayCounter <= endMonthDays ):
            ws['{0}{1}'.format(col,3)].style = style_day
            ws['{0}{1}'.format(col,2)].style = style_month
            if day.strftime('%w') == "0" or day.strftime('%w') == "6":
                ws['{0}{1}'.format(col,4)].style = style_weekend
            else:
                ws['{0}{1}'.format(col,4)].style = style_weekday
        else:
            ws['{0}{1}'.format(col,3)].style = style_day_inactive
            ws['{0}{1}'.format(col,2)].style = style_month_inactive
            if day.strftime('%w') == "0" or day.strftime('%w') == "6":
                ws['{0}{1}'.format(col,4)].style = style_weekend_inactive
            else:
                ws['{0}{1}'.format(col,4)].style = style_weekday_inactive

        ws.column_dimensions[col].width = 5

    # merge month header 
    if (endHeaderDays > startHeaderDays):
        ws.merge_cells('{0}{1}:{2}{3}'.format(get_column_letter(offset_cols + startHeaderDays),2,get_column_letter(offset_cols + endHeaderDays),2))
    
    ws.merge_cells('{0}{1}:{2}{3}'.format(get_column_letter(offset_cols + endHeaderDays + 1),2,get_column_letter(offset_cols + endMonthDays),2))
    # 1 + date.weekday()+calendar.monthrange(date.year,date.month)[1]
    if (endMonthDays < endTailDates):
        ws.merge_cells('{0}{1}:{2}{3}'.format(get_column_letter(offset_cols + endMonthDays + 1),2,get_column_letter(offset_cols + endTailDates),2))


    # Team member lines


    ws['{}{}'.format(get_column_letter(team_offset_cols),team_offset_rows-1)]='Team'
    ws['{}{}'.format(get_column_letter(team_offset_cols),team_offset_rows-1)].style = style_team_header
    ws.merge_cells('{0}{1}:{2}{3}'.format(get_column_letter(team_offset_cols),team_offset_rows-1,get_column_letter(team_offset_cols),team_offset_rows))
    ws.row_dimensions[team_offset_rows].height = 4

    # add empty columns
    for i in range(1,team_offset_cols):
        ws.column_dimensions[get_column_letter(i)].width = 1
    # add empty columns
    for i in range(team_offset_cols+1,offset_cols+1):
        ws.column_dimensions[get_column_letter(i)].width = 1

    counter = 0
    for teamMember in teamMembers:
        counter += 1
        column_nr = team_offset_cols

        row = team_offset_rows + counter
        col = get_column_letter(column_nr)
        
        ws['{0}{1}'.format(col,row)]=teamMember[0]
        if (teamMember[1] is None or teamMember[1] <= date) and (teamMember[2] is None or teamMember[2] > date):
            if (counter % 2) == 0:
                ws['{0}{1}'.format(col,row)].style = style_team_even
            else:
                ws['{0}{1}'.format(col,row)].style = style_team_odd
        else:
            ws['{0}{1}'.format(col,row)].style = style_team_inactive

        


        column_nr = offset_cols
        dayCounter = 0

        days = cal.itermonthdates(date.year, date.month)

        for day in days:
            column_nr += 1
            dayCounter += 1

            col = get_column_letter(column_nr)

            if (teamMember[1] is None or teamMember[1] <= date) and (teamMember[2] is None or teamMember[2] > date):
                if ( dayCounter > endHeaderDays and dayCounter <= endMonthDays  ):
                    # weekend days are non-working days and greyed out as such
                    if day.strftime('%w') == "0" or day.strftime('%w') == "6":
                        ws['{0}{1}'.format(col,row)]=""
                        ws['{0}{1}'.format(col,row)].style=style_weekend
                    else:
                        # check if sprint flip day - style_flip_day
                        if (day in flip_days):
                            ws['{0}{1}'.format(col,row)].style=style_flip_day
                        else:
                            if (teamMember[2 + int(day.strftime('%w'))]==0):
                                ws['{0}{1}'.format(col,row)].style=style_weekend
                            else:
                                if (counter % 2) == 0:
                                    ws['{0}{1}'.format(col,row)].style=style_workday_even
                                else:
                                    ws['{0}{1}'.format(col,row)].style=style_workday_odd

                else:
                    # inactive days are inactive and copy the value from the previous month 
                    

                    # inactive days are locked
                    if day.strftime('%w') == "0" or day.strftime('%w') == "6":
                        ws['{0}{1}'.format(col,row)]="x"
                        ws['{0}{1}'.format(col,row)].style=style_weekend_inactive
                    else:
                        ws['{0}{1}'.format(col,row)].style=style_workday_inactive
            else:
                if (day in flip_days):
                    ws['{0}{1}'.format(col,row)].style=style_flip_day
                else:
                    ws['{0}{1}'.format(col,row)].style=style_team_workday_inactive


    # summary
    column_nr = offset_cols
    dayCounter = 0

    days = cal.itermonthdates(date.year, date.month)

    offset_footer = 3
    row = team_offset_rows + counter + offset_footer

    for day in days:
        column_nr += 1
        dayCounter += 1

        col = get_column_letter(column_nr)

        if day.strftime('%w') == "0" or day.strftime('%w') == "6":
            ws['{0}{1}'.format(col,row)].style=style_weekend
        else:
            if (day in flip_days):
                ws['{0}{1}'.format(col,row)].style=style_flip_day
            else:
                if ( dayCounter > endHeaderDays and dayCounter <= endMonthDays  ):
                    ws['{0}{1}'.format(col,row)]="=SUM({}{}:{}{})".format(col,team_offset_rows,col,team_offset_rows+len(teamMembers))
                    ws['{0}{1}'.format(col,row)].style=style_day_sum
                else:
                    ws['{0}{1}'.format(col,row)]=""
                    ws['{0}{1}'.format(col,row)].style=style_weekend
        




# Save the file
filename = "output\plan.xlsx"
wb.save(filename)

# open the Excel file
os.system("{}".format(filename))