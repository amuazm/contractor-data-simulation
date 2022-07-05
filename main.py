import shutil
from openpyxl import Workbook, load_workbook
import datetime
import random

import openpyxl

src = "./Budget Dataset Modified.xlsx"
dest = "./Output/Backup.xlsx"
shutil.copyfile(src, dest)

data = load_workbook("./Output/Backup.xlsx")
result = Workbook()

data_ws = data["Project Details"]
project_ids = data_ws["A"][1:]
project_budgets = data_ws["L"][1:]

result_ws = result.active
result_ws.title = "Project Details"

# Headers
result_ws.append(["Project ID", "Start Date", "Duration (Months)", "Cost per Month", "Total Cost", "Contract Pay"])

date_style = openpyxl.styles.NamedStyle(name="date_style", number_format="dd/mm/yyyy")
i = 0
for row in result_ws.iter_rows(min_row=2, max_row=len(project_ids), max_col=6):
    # Project IDs
    row[0].value = project_ids[i].value

    # Start Date
    row[1].style = date_style
    start = datetime.datetime.strptime("01/01/2021", "%d/%m/%Y").date()
    end = datetime.date.today()
    delta = end - start
    r = random.randrange(delta.days)
    new = start + datetime.timedelta(days=r)
    row[1].value = new

    # Duration (Months)
    row[2].value = random.randrange(12, 25)

    # Cost per Month
    row[3].style = "Currency"
    row[3].value = project_budgets[i].value / row[2].value

    # Total Cost
    row[4].style = "Currency"
    row[4].value = project_budgets[i].value

    # Contract Pay
    row[5].style = "Currency"
    row[5].value = project_budgets[i].value * random.randint(1020, 1080)/1000

    i += 1

data.close()
result.save("./Output/Result.xlsx")