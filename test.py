# Imports
from types import NoneType
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

# Iterators
i = 2
j = 5
# Empty lists
bricks = list()
req_fields = list()
opt_fields = list()
req_fields_vals = list()
opt_fields_vals = list()
# Sheets
dirty_wb = openpyxl.load_workbook('All SKUs 21.06.xlsx')
dirty_sheets = dirty_wb.sheetnames
dirty_sheet = dirty_wb[dirty_sheets[0]]
glossary_wb = openpyxl.load_workbook('202205-GS1-Datamodel-DHZTD-3.1.19_EN.xlsx')
glossary_sheets = glossary_wb.sheetnames
glossary_sheet = glossary_wb[glossary_sheets[2]]
# Temp
tmp_g_brick = 0
# Find unique categories
for row in dirty_sheet:
    brick = dirty_sheet.cell(column = 5, row = i).value
    i += 1
    if(brick is None or brick is NoneType):
        continue
    elif (int(brick) in bricks):
        continue
    bricks.append(int(brick))

for row in glossary_sheet:
    # Check g_brick belongs to our list
    if((glossary_sheet.cell(column=1,row = j).value) in bricks):
        # Check brick has no more atr(0/1) a.k.a. glossary has jumped to a new brick
        if(glossary_sheet.cell(column=1,row = j).value != tmp_g_brick
        and tmp_g_brick != 0):
            clean_wb = Workbook()
            clean_ws = clean_wb.active
            # Write tab with req and opt for one brick
            clean_ws.append(req_fields)
            clean_ws.append(opt_fields)
            for cell in clean_ws['A1:'+chr(len(req_fields)+65)+'1'][0]:
                cell.fill = PatternFill("solid", start_color="ed5587")
                cell.font = Font(bold=True)
            for cell in clean_ws['A2:CC2'][0]:
                cell.fill = PatternFill("solid", start_color="97ed55")
                cell.font = Font(italic=True)
            clean_wb.save(str(tmp_g_brick)+".xlsx")
            # Clean lists for new brick
            req_fields = list()
            opt_fields = list()
            clean_wb.close()

        # This should happen first... Go to Col C
        if(glossary_sheet.cell(column=3,row = j).value == 1):
            # Attach (req) atr to list
            req_fields.append((glossary_sheet.cell(column=6,row = j).value).upper())
        elif(glossary_sheet.cell(column=3,row = j).value == 0):
            # Attach (opt) atr to list
            opt_fields.append((glossary_sheet.cell(column=6,row = j).value).lower())

        tmp_g_brick = glossary_sheet.cell(column=1,row = j).value
    j += 1