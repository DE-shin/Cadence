import win32com.client
import pythoncom
import pandas as pd
import base64
import os
from bs4 import BeautifulSoup
from PIL import Image
from io import BytesIO, StringIO
import logging

logger = logging.getLogger(__name__)

def post_sim(HTML_FILE_PATH, REPORT_FOLDER_PATH):

    # Read report.htm
    with open(os.path.normpath(HTML_FILE_PATH), "r", encoding="utf-8") as f:
        html_content = f.read()

    report = BeautifulSoup(html_content, "html.parser")

    # Parsing Table Datas
    table_setups = report.select("ElectricalSetup table")
    table_setup = str(table_setups[2])
    df_setup = pd.read_html(StringIO(table_setup), header=0)[0]

    table_results = report.select("ElectricalResultsTable table")
    table_result = str(table_results[2])
    df_result = pd.read_html(StringIO(table_result), header=0)[0]

    # Processing Tables
    common_columns = df_setup.columns.intersection(df_result.columns)
    df_merged = pd.merge(df_setup, df_result, on=list(common_columns), how="inner")

    df_new = df_merged.iloc[:, [1, 0, 2, 5, 14, 15, 16]].reset_index(drop=True)
    df_new[df_new.columns[0]] = df_new[df_new.columns[0]].str.replace("SINK_", "")
    df_new[df_new.columns[2]] = df_new[df_new.columns[2]].astype(str).str.split("-").str[0]

    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Add()
    ws = wb.Worksheets(1)

    for r_idx, row in enumerate(df_new.itertuples(index=False, name=None), 2):
        for c_idx, value in enumerate(row, 1):
            ws.Cells(r_idx, c_idx).Value = value

    for c_idx, col_name in enumerate(df_new.columns, 1):
        ws.Cells(1, c_idx).Value = col_name

    font_name = "현대하모니 L"
    font_size = 16

    for cell in ws.UsedRange:
        cell.Font.Name = font_name
        cell.Font.Size = font_size
        cell.Borders.LineStyle = 1

    wb.SaveAs(os.path.normpath(os.path.join(REPORT_FOLDER_PATH, "report.xlsx")))
    wb.Close(False)
    excel.Quit()
    del excel
    pythoncom.CoUninitialize()

    # Parsing Images
    layers = ["ImageLayoutTop", "ImageLayoutBottom"]
    for layer in layers:
        image_element = report.select_one(f"#{layer} img")
        if image_element:
            mage_data_url = img["src"]
            base64_data = image_data_url.split(",")[1]
            image_data = base64.b64decode(base64_data)
            image = Image.open(BytesIO(image_data))
            image.save(os.path.normpath(os.path.join(REPORT_FOLDER_PATH, f"{layer}.png")))

    image_plot = report.select("#DistributionPlot p img")
    for idx, img in enumerate(image_plot):
        image_data_url = img["src"]
        base64_data = image_data_url.split(",")[1]
        image_data = base64.b64decode(base64_data)
        image = Image.open(BytesIO(image_data))
        image.save(os.path.normpath(os.path.join(REPORT_FOLDER_PATH, f"Layer_{idx+1}.png")))

    return None