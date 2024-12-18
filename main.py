import xlwings as xl
import pandas as pd
from sim_pre import net_classify_power

file_path = ""
prj_path = ""

wb_user = xl.Book(file_path)

sheets_user = wb_user.sheets
dfs_user = {sheet.name: sheet.used_range.options(pd.DataFrame, header=1, index=False).value for sheet in sheets_user}

wb_user.close()

net_classify_power(dfs=dfs_user, prj_path=prj_path)
