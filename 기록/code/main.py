import openpyxl as oxl
from openpyxl import load_workbook
import numpy as np
from datetime import datetime

person = "김다은"
content_colNum = 3
tag_colNum = 5

# create workbook
wb = oxl.Workbook()

# open xlsx file named "Data.xlsx"
lwb = oxl.load_workbook("Data.xlsx")

# get the sheet named "전체"
sheet = lwb["전체"]

# get the data from the sheet A to G which has B column has value of "박지훈"
data = []
for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=7):
    if row[1].value == person:
        data.append([cell.value for cell in row])

# add element of str - "(" + 'value of col named "tag"' + ") " + 'value of col named "content"' to data
for i in range(len(data)):
    data[i].append("(" + str(data[i][tag_colNum]) + ") " + str(data[i][content_colNum]))

# Print data
for i in range(len(data)):
    print(data[i])

'''
# create sheet named "timeline_" + person
sheet = lwb.create_sheet("timeline_" + person)
# write the data to the new sheet
for i in range(len(data)):
    for j in range(len(data[i])):
        sheet.cell(row=i+1, column=j+1, value=data[i][j])

# save the workbook
lwb.save("Data.xlsx")

'''