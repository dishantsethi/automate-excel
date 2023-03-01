from openpyxl import load_workbook
import os
from config import INPUT_DIR, OUTPUT_DIR
from utils import get_sheets_from_input_dir
from functions import insert_second_row_in_summary, page_setup_for_each_sheet
import time

for file in get_sheets_from_input_dir():
    try:
        wb = load_workbook(filename = f"{INPUT_DIR}/{file}")
        print("--------------------------------------------")
        print(f"File name {file}")
        tic = time.time()
        insert_second_row_in_summary(wb, file)
        page_setup_for_each_sheet(wb, file)
        toc = time.time()
        print(f"Time Take: {toc-tic} seconds")
        print("--------------------------------------------")
    except Exception as e:
        print(e)
    else:
        os.rename(f"{INPUT_DIR}/{file}", f"{OUTPUT_DIR}/{file}")
        