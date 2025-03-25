import logging
from presim import PdcPresim, PsiPresim
from postsim import *

def setup_logger():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.FileHandler("Cadence.log", mode="w", encoding="utf-8"),
            logging.StreamHandler()
        ]
    )

    return None

if __name__ == "__main__":
    setup_logger()

    # 유저 변수
    GND_NAME = "GND"
    ETL_PDC_FILE_PATH = ""
    ETL_PSI_FILE_PATH = ""
    TCL_FOLDER_PATH = ""
    REPORT_FILE_PATH = ""
    REPORT_FOLDER_PATH = ""

    # 함수 호출
    PdcPresim(GND_NAME, ETL_PDC_FILE_PATH)
    PsiPresim(GND_NAME, ETL_PSI_FILE_PATH)
    pdc_postsim(REPORT_FILE_PATH, REPORT_FOLDER_PATH)