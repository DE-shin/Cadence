import pandas as pd
import xlwings as xw
import base64
import os
from bs4 import BeautifulSoup
from PIL import Image
from io import BytesIO, StringIO
import logging

logger = logging.getLogger(__name__)

def post_sim(HTML_FILE_PATH, REPORT_FOLDER_PATH):
    REPORT_PIC_FOLDER_PATH = os.path.normpath(os.path.join(REPORT_FOLDER_PATH, "pics"))
    REPORT_EXCEL_FILE_PATH = os.path.normpath(os.path.join(REPORT_FOLDER_PATH, f"{os.path.basename(HTML_FILE_PATH).split(".")[0]}_PDC_Result.xlsx"))

    logger.info("Reading html report...")
    try:
        with open(os.path.normpath(HTML_FILE_PATH), "r", encoding="utf-8") as f:
            html_content = f.read()
        report = BeautifulSoup(html_content, "html.parser")
    except Exception as e:
        logger.error(f"Error reading HTML file: {e}")
        return None
    logger.info("Reading Complete!\n")

    try:
        table_setups = report.select("#ElectricalSetup table")
        table_setup = str(table_setups[2])
        df_setup = pd.read_html(StringIO(table_setup), header=0)[0]

        table_results = report.select("#ElectricalResultsTable table")
        table_result = str(table_results[2])
        df_result = pd.read_html(StringIO(table_result), header=0)[0]

        common_columns = df_setup.columns.intersection(df_result.columns)
        df_merged = pd.merge(df_setup, df_result, on=list(common_columns), how="inner")

        df_new = df_merged.iloc[:, [1, 0, 2, 5, 14, 15, 16]].reset_index(drop=True)
        df_new[df_new.columns[0]] = df_new[df_new.columns[0]].str.replace("SINK_", "")
        df_new[df_new.columns[2]] = df_new[df_new.columns[2]].astype(str).str.split("-").str[0]
    except Exception as e:
        logger.error(f"Error processing tables: {e}")
        return None

    logger.info("Generating Excel Report...")
    try:
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.add()
        ws = wb.sheets["Sheet1"]

        ws.range("A1").value = [df_new.columns.tolist()] + df_new.values.tolist()

        last_row = ws.range("A1").end("down").row
        last_col = ws.range("A1").end("right").column

        data_range = ws.range(f"A1:{ws.range((last_row, last_col)).address}")
        header_range = ws.range(f"A1:{ws.range((1, last_col)).address}")
        pass_fail_range = ws.range(f"F2:F{last_row}")

        header_range.color = (204, 255, 153)
        data_range.api.Borders.Weight = 2
        data_range.api.Font.Name = "현대하모니 L"
        data_range.api.Font.Size = 16

        for cell in pass_fail_range:
            if cell.value == "Fail":
                row_range = ws.range(f"A{cell.row}:{ws.range((cell.row, last_col)).address}")
                row_range.color = (255, 204, 204)
                row_range.api.Font.Color = -16776961

        prev_value = None
        merge_start = 2
        for row in range(2, last_row + 1):
            current_value = ws.range(f"A{row}").value
            if current_value == prev_value:
                ws.range(f"A{merge_start}:A{row}").merge()
            else:
                merge_start = row
            prev_value = current_value
        
        ws.range(f"A1:{ws.range((1, last_col)).address}").api.EntireColumn.AutoFit()
        logger.info("Generating Complete!\n")

        os.makedirs(REPORT_FOLDER_PATH, exist_ok=True)
        wb.save(REPORT_EXCEL_FILE_PATH)
        logger.info(f"Save : {REPORT_EXCEL_FILE_PATH}\n")
    except Exception as e:
        logger.error(f"Error saving Excel file: {e}")
    finally:
        wb.close()
        app.quit()

    logger.info("Saving Images....")
    try:
        layers = ["ImageLayoutTop", "ImageLayoutBottom"]
        for layer in layers:
            image_element = report.select_one(f"#{layer} img")
            if image_element and "src" in image_element.attrs:
                image_data_url = image_element["src"]
                base64_data = image_data_url.split(",")[1]
                image_data = base64.b64decode(base64_data)
                image = Image.open(BytesIO(image_data))

                os.makedirs(REPORT_PIC_FOLDER_PATH, exist_ok=True)
                image.save(os.path.normpath(os.path.join(REPORT_FOLDER_PATH, f"{layer}.png")))
        logging.info(f"Saved : {REPORT_PIC_FOLDER_PATH}_{layer}.png")
    except Exception as e:
        logger.error(f"Error saving Top/Bottom images: {e}")

    logger.info("\nSaving Images....")
    try:
        image_plot = report.select("#DistributionPlot p img")
        for idx, img in enumerate(image_plot):
            if "src" in img.attrs:
                image_data_url = img["src"]
                base64_data = image_data_url.split(",")[1]
                image_data = base64.b64decode(base64_data)
                image = Image.open(BytesIO(image_data))

                os.makedirs(REPORT_PIC_FOLDER_PATH, exist_ok=True)
                image.save(os.path.normpath(os.path.join(REPORT_FOLDER_PATH, f"Layer_{idx+1}.png")))
        logging.info(f"Saved : Layer 1 ~ {idx+1} Images\n")
    except Exception as e:
        logger.error(f"Error saving Layer images: {e}")

    return None
