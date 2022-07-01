import shutil
from openpyxl import Workbook, load_workbook

src = "./Budget Dataset Modified.xlsx"
dest = "./Output/Backup.xlsx"
shutil.copyfile(src, dest)

data = load_workbook("./Output/Backup.xlsx")
result = Workbook()

ws = data["Project Details"]
projectIds = ws["A"][1:]

data.close()
result.save("./Output/Result.xlsx")