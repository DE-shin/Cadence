import os
import inspect
import base64
import logging
import pandas as pd
import xlwings as xw
from bs4 import BeautifulSoup
from PIL import Image
from io import BytesIO, StringIO

logger = logging.getLoggger(__name__)

# PowerDC
class pdc_postsim:
    def __init__(self, REPORT_FILE_PATH, OUTPUT_FOLDER_PATH):
        # 상수
        self.report_file_path = os.path.normpath(REPORT_FILE_PATH)
        self.output_folder_path = os.path.normpath(OUTPUT_FOLDER_PATH)
        self.output_pic_folder_path = os.path.normpath(os.path.join(self.output_folder_path, "pics"))
        self.output_excel_file_path = os.path.normpath(os.path.join(self.output_folder_path, f"{os.path.basename(self.report_file_path).split(".")[0]}_PDC_Result.xlsx"))

        # 변수
        self.report = None

        # 함수
        self.initialize()
        self.extract_excel()
        self.extract_images()

    def initialize(self):
        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 실행 중")

        os.makedirs(self.output_folder_path, exist_ok=True)
        os.makedirs(self.output_pic_folder_path, exist_ok=True)

        with open(self.report_file_path, "r", encoding="utf-8") as f:
            raw_report = f.read()
        self.report = BeautifulSoup(raw_report, "html.parser")

        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 완료")
        return None

    def extract_excel(self):
        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 실행 중")

        # Table 데이터 추출
        table_setups = self.report.select("#ElectricalSetup table")
        table_setup = str(table_setups[2])
        df_setup = pd.read_html(StringIO(table_setup), header=0)[0]

        table_results = self.report.select("#ElectricalResultsTable table")
        table_result = str(table_results[2])
        df_result = pd.read_html(StringIO(table_result), header=0)[0]

        common_columns = df_setup.columns.intersection(df_result.columns)
        df_merged = pd.merge(df_setup, df_result, on=list(common_columns), how="inner")

        df_new = df_merged.iloc[:, [1, 0, 2, 5, 14, 15, 16]].reset_index(drop=True)
        df_new[df_new.columns[0]] = df_new[df_new.columns[0]].str.replace("SINK_", "")
        df_new[df_new.columns[2]] = df_new[df_new.columns[2]].astype(str).str.split("-").str[0]

        # 엑셀 파일로 내보내기
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
        data_range.api.Font.Name = "현대하모니 M"
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
        wb.save(self.output_excel_file_path)

        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 완료")
        return None
    
    def extract_images(self):
        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 실행 중")

        # Top/Bottem Layer 사진 추출
        layers = ["ImageLayoutTop", "ImageLayoutBottom"]

        for layer in layers:
            image_element = self.report.select_one(f"#{layer} img")
            if image_element and "src" in image_element.attrs:
                image_data_url = image_element["src"]
                base64_data = image_data_url.split(",")[1]
                image_data = base64.b64decode(base64_data)
                image = Image.open(BytesIO(image_data))

                image.save(os.path.join(self.output_pic_folder_path, f"{layer}.png"))

        # Distribution Plot 사진 추출
        image_plot = self.report.select("#DistributionPlot p img")

        for idx, img in enumerate(image_plot):
            if "src" in img.attrs:
                image_data_url = img["src"]
                base64_data = image_data_url.split(",")[1]
                image_data = base64.b64decode(base64_data)
                image = Image.open(BytesIO(image_data))

                image.save(os.path.join(self.output_pic_folder_path, f"Layer_{idx+1}.png"))

        logger.info(f"{self.__class__.__name__} : {inspect.currentframe().f_code.co_name} 완료")
        return None