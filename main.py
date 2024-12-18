import xlwings as xl
import pandas as pd
from sim_pre import classify_power_net, add_VRM_SINK, add_DCR

file_path = ""
prj_path = ""

wb_user = xl.Book(file_path)

sheets_user = wb_user.sheets
dfs_user = {sheet.name: sheet.used_range.options(pd.DataFrame, header=1, index=False).value for sheet in sheets_user}

wb_user.close()

classify_power_net(dfs=dfs_user, prj_path=prj_path)
add_VRM_SINK(dfs=dfs_user, prj_path=prj_path)
add_DCR(dfs=dfs_user, prj_path=prj_path)
