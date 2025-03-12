import logging
import time
from tcl import pre_sim
from report import post_sim

def kst_time(*args):
    return time.locatime(time.mktime(time.gmtime()) + 9*3600)

# Logger Settings
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y/%m/%d %H:%M:%S",
    filename="cadence.log",
    filemode="w"
    encoding="utf-8",
    force=True
)

console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_formatter = logging.Formatter(
    "%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y/%m/%d %H:%M:%S"
)
console_formatter.converter = kst_time
console_handler.setFormatter(console_formatter)
logging.getLogger().addHandler(console_handler)

# Paths
ETL_FILE_PATH = ""
TCL_FOLDER_PATH = ""
HTML_FILE_PATH = ""
REPORT_FOLDER_PATH = ""

# Execute
pre_sim(ETL_FILE_PATH, TCL_FOLDER_PATH)
post_sim(HTML_FILE_PATH, REPORT_FOLDER_PATH)
