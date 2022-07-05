from types import NoneType
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
from datetime import datetime

# Iterators
i = 2
j = 5
# Empty lists
bricks = list()
req_fields = list()
opt_fields = list()
req_fields_vals = dict()
opt_fields_vals = dict()
column_widths = []
gs1_ids = dict()
# Sheets
dirty_wb = openpyxl.load_workbook('required_files\All SKUs 21.06.xlsx')
dirty_sheets = dirty_wb.sheetnames
dirty_sheet = dirty_wb[dirty_sheets[0]]
glossary_wb = openpyxl.load_workbook('required_files\glossary.xlsx')
glossary_sheets = glossary_wb.sheetnames
glossary_sheet = glossary_wb[glossary_sheets[2]]
# Temp
tmp_g_brick = 0
# Find unique categories
for row in dirty_sheet:
    id = dirty_sheet.cell(column=1, row=i).value
    brick = dirty_sheet.cell(column=5, row=i).value
    gs1_ids[id] = [i, brick]
    i += 1
    if(brick is None or brick is NoneType):
        continue
    elif (int(brick) in bricks):
        continue
    bricks.append(int(brick))

for row in glossary_sheet:
    # Check g_brick belongs to our list
    if((glossary_sheet.cell(column=1, row=j).value) in bricks):
        # Check brick has no more atr(0/1) a.k.a. glossary has jumped to a new brick
        if(glossary_sheet.cell(column=1, row=j).value != tmp_g_brick
           and tmp_g_brick != 0):
            clean_wb = Workbook()
            clean_ws = clean_wb.active
            # Write tab with req and opt for one brick
            clean_ws.append(["GS1 item number (GTIN):", "", "BRICK CAT:", tmp_g_brick, "v0.1",
                            "Script last executed:", datetime.today().strftime('%Y-%m-%d %H:%M:%S')])
            clean_ws.append(req_fields)
            clean_ws.append(req_fields_vals)
            clean_ws.append(opt_fields)
            clean_ws.append(opt_fields_vals)
            for cell in clean_ws['A2:'+chr(len(req_fields)+65)+'2'][0]:
                cell.fill = PatternFill("solid", start_color="ed5587")
                cell.font = Font(bold=True)
            for cell in clean_ws['A4:BR4'][0]:
                cell.fill = PatternFill("solid", start_color="97ed55")
                cell.font = Font(italic=True)
            for row in clean_ws:
                for i, cell in enumerate(row):
                    if len(column_widths) > i:
                        if len(str(cell)) > column_widths[i]:
                            column_widths[i] = len(cell)
                    else:
                        column_widths += [len(str(cell))]

            # ,1 to start at 1
            for i, column_width in enumerate(column_widths, 1):
                clean_ws.column_dimensions[get_column_letter(
                    i)].width = column_width
            clean_wb.save("results\\"+str(tmp_g_brick)+".xlsx")
            # Clean lists for new brick
            req_fields = list()
            opt_fields = list()
            clean_wb.close()

        # This should happen first... Go to Col C
        if(glossary_sheet.cell(column=3, row=j).value == 1):
            # Attach (req) atr to list
            req_fields.append(
                (glossary_sheet.cell(column=6, row=j).value).upper())
            for col in dirty_sheet["A1:AWK1"]:
                for k, atr in enumerate(col):
                    for req_field in req_fields:
                        if req_field.startswith(atr.value.upper()):
                            for gs1_i in gs1_ids:
                                req_fields_vals[gs1_ids[gs1_i][0]] = str(
                                    dirty_sheet.cell(column=k, row=gs1_ids[gs1_i][0]).value)
                    for opt_field in opt_fields:
                        if opt_field.startswith(atr.value.lower()):
                            for gs1_i in gs1_ids:
                                opt_fields_vals[gs1_ids[gs1_i][0]] = str(
                                    dirty_sheet.cell(column=k, row=gs1_ids[gs1_i][0]).value)
        elif(glossary_sheet.cell(column=3, row=j).value == 0):
            # Attach (opt) atr to list
            opt_fields.append(
                (glossary_sheet.cell(column=6, row=j).value).lower())

        tmp_g_brick = glossary_sheet.cell(column=1, row=j).value
    j += 1
