from types import NoneType
import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

i = 2
j = 5
k = 1
bricks = list()
req_fields = list()
opt_fields = list()

dirty_wb = openpyxl.load_workbook('All SKUs 21.06.xlsx')
dirty_sheets = dirty_wb.sheetnames
dirty_sheet = dirty_wb[dirty_sheets[0]]

glossary_wb = openpyxl.load_workbook('202205-GS1-Datamodel-DHZTD-3.1.19_EN.xlsx')
glossary_sheets = glossary_wb.sheetnames
glossary_sheet = glossary_wb[glossary_sheets[2]]
prev_glossary_brick = 0

clean_wb = Workbook()
clean_ws = clean_wb.active

for row in dirty_sheet:
    brick = dirty_sheet.cell(column = 5, row = i).value
    i += 1
    if(brick in bricks or brick is None or brick is NoneType):
        continue
    bricks.append(brick)

for row in glossary_sheet:
    # Glossary brick is of importance
    if((glossary_sheet.cell(column=1,row = j).value) in bricks):
        # New brick, new fields
        if(glossary_sheet.cell(column=3,row = j).value != prev_glossary_brick and prev_glossary_brick != 0):
            # Write tab with req and opt for one brick
            clean_ws.append([prev_glossary_brick])
            for field in req_fields:
                clean_ws.append(field)
                k+=1
            for field in opt_fields:
                clean_ws.append(field)
                k+=1
            tab = Table(displayName="Table"+str(k), ref="A1:E"+str(k))
            style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                       showLastColumn=False, showRowStripes=True, showColumnStripes=True)
            tab.tableStyleInfo = style
            clean_ws.add_table(tab)
            clean_wb.save(str(glossary_sheet.cell(column=1,row = j).value)+".xlsx")
            # Clean the lists
            req_fields = list()
            opt_fields = list()

        # Go to Col C
        if(glossary_sheet.cell(column=3,row = j).value == 1):
            # Attach (req) atr to list
            req_fields.append(glossary_sheet.cell(column=6,row = j).value)
        elif(glossary_sheet.cell(column=3,row = j).value == 0):
            # Attach (opt) atr to list
            opt_fields.append(glossary_sheet.cell(column=6,row = j).value)
        prev_glossary_brick = glossary_sheet.cell(column=1,row = j).value
    j += 1