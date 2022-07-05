import shutil
from openpyxl import Workbook, load_workbook
import datetime
import random
import numpy as np
from dateutil import relativedelta

import openpyxl

src = "./Budget Dataset Modified.xlsx"
dest = "./Output/Backup.xlsx"
shutil.copyfile(src, dest)

data = load_workbook("./Output/Backup.xlsx")
result = Workbook()

data_ws = data["Project Details"]
project_ids = data_ws["A"][1:]
project_ids = [i.value for i in project_ids]
project_budgets = data_ws["L"][1:]
project_budgets = [i.value for i in project_budgets]

# Budget
if True:
    result_ws = result.active
    result_ws.title = "Budget"

    # Headers
    result_ws.append(["Project ID", "Start Date", "Duration (Months)", "Cost per Month", "Total Cost", "Contract Pay"])

    date_style = openpyxl.styles.NamedStyle(name="date_style", number_format="dd/mm/yyyy")
    i = 0
    for row in result_ws.iter_rows(min_row=2, max_row=len(project_ids) + 1, max_col=6):
        # Project IDs
        row[0].value = project_ids[i]

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
        row[3].value = project_budgets[i] / row[2].value

        # Total Cost
        row[4].style = "Currency"
        row[4].value = project_budgets[i]

        # Contract Pay
        row[5].style = "Currency"
        row[5].value = project_budgets[i] * random.randint(1020, 1080)/1000

        i += 1

# Categories
if True:
    result.create_sheet("Categories")
    result_ws = result["Categories"]

    #Headers
    result_ws.append(["Project ID", "Category", "Budget"])

    categories = [
        'Direct Labour',
        'Supplied Labour',
        'Sub-contractor',
        'Other Materials',
        'Small Tools & Safety Item',
        'Other Consumable',
        'Transportation',
        'Repair & Maintenance',
        'Site Office Expense',
        'Food, Refreshment & Entertainment',
        'Travelling & Vehicles',
        'Main Steel Materials',
        'Stainless Steel Materials',
        'Aluminium Materials',
        'Equipment',
        'Transportation',
        'Supervision',
        'Insurance'
        ]

    result_ws = result["Budget"]
    budgets = result_ws["E"][1:]
    budgets = [i.value for i in budgets]
    result_ws = result["Categories"]
    i2 = 0
    i4 = 1
    for i in project_ids:
        n, k = budgets[i2], len(categories)
        vals = np.random.default_rng().dirichlet(np.ones(k), size=1)
        k_nums = [round(v) for v in vals[0]*n]
        i3 = 0
        for category in categories:
            i4 += 1
            result_ws.append([i, category, k_nums[i3]])
            result_ws[f"C{i4}"].style = "Currency"
            i3 += 1
        i2 += 1

# Cash Outflow
if True:
    result.create_sheet("Cash Outflow")
    result_ws = result["Cash Outflow"]

    # Headers
    result_ws.append(["Project ID", "Month", "Category", "Amount"])

    start_dates = result["Budget"]["B"][1:]
    start_dates = [i.value for i in start_dates]
    duration = result["Budget"]["C"][1:]
    duration = [i.value for i in duration]
    budgets_by_cat = result["Categories"]["C"][1:]
    budgets_by_cat = [i.value for i in budgets_by_cat]

    i2 = 0
    i5 = 1
    for i in project_ids:
        months_completed = random.randrange(6, duration[i2])
        for i3 in range(months_completed):
            i4 = 0
            for category in categories:
                i5 += 1
                the_date = start_dates[i2]
                the_date += relativedelta.relativedelta(months=i3)
                result_ws.append([i, the_date, category, budgets_by_cat[i4]/duration[i2]*(random.randint(900, 1100)/1000)])
                result_ws[f"D{i5}"].style = "Currency"
                i4 += 1
        i2 += 1

# Reports
if False:
    result.create_sheet("Reports")
    result_ws = result["Reports"]

    # Headers
    result_ws.append(["Project ID", "Report Date", "Completion", "Incurred Cost"])

    for i in project_ids:
        result_ws.append([i])

data.close()
result.save("./Output/Result.xlsx")