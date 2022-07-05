import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
from datetime import datetime
import os.path
import os

print("Loading files")
allsku = openpyxl.load_workbook('All_SKUs_21.06.xlsx')
allsku = allsku[allsku.sheetnames[0]]
print("Loaded all sku")

glossary = openpyxl.load_workbook('glossary.xlsx')
glossary = glossary[glossary.sheetnames[2]]
glossary.delete_rows(1, 4)
print("Loaded glossary")


headers = [cell.value for cell in allsku[1]] # all product column names

print(headers)
#           category name   mandatory


categories = {}
for row in glossary:
    categoryid = str(row[0].value)
    if not categories.get(categoryid, None):
        categories[categoryid] = [(row[5].value, row[2].value)]
    else:
        categories[categoryid].extend([(row[5].value, row[2].value)])

for i, row in enumerate(allsku):
    if i == 0: # skip headers
        continue

    categoryid = str(row[4].value)
    if not os.path.exists(categoryid):
        os.mkdir(categoryid)


    clean_wb = Workbook()
    clean_ws = clean_wb.active

    local_req = headers[::]
    req = [column for column, priority in categories[categoryid] if priority == 1]
    local_req = req[::]
    req_vals = []
    for i, cell in enumerate(row):
        for required in local_req:
            if headers[i].lower().startswith(required.lower()):
                to_remove = required
                break
        else:
            continue
        local_req.remove(to_remove)
        req_vals.append(cell.value)

    # req_vals = [cell.value for i, cell in enumerate(row) if any([required.lower().startswith(headers[i].lower()) for required in req])]

    opt = [column for column, priority in categories[categoryid] if priority == 0]
    opt_vals = [cell.value for i, cell in enumerate(row) if headers[i] in opt]

    print(len(req))
    print(len(req_vals))
    print()
    # print(opt)
    # print(opt_vals)