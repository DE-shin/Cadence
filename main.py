import logging
from tcl import pre_sim
from report import post_sim

# Logger Settings
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    filename="tcl.log",
    filemode="w"
)

# Paths
ETL_FILE_PATH = ""
TCL_FOLDER_PATH = ""
HTML_FILE_PATH = ""
REPORT_FOLDER_PATH = ""

pre_sim(ETL_FILE_PATH, TCL_FOLDER_PATH)
post_sim(HTML_FILE_PATH, REPORT_FOLDER_PATH)
