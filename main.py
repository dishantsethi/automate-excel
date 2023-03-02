from openpyxl import load_workbook
import os
from config import INPUT_DIR
from utils import get_sheets_from_input_dir, move_files_to_output_folder
from functions import insert_second_row_in_summary, page_setup_for_each_sheet
import time

for file in get_sheets_from_input_dir():
    filename = os.path.join(INPUT_DIR, file)
    try:
        wb = load_workbook(filename = filename)
        print("--------------------------------------------")
        print(f"File name {file}")
        tic = time.time()
        insert_second_row_in_summary(wb, file)
        page_setup_for_each_sheet(wb, file)
        toc = time.time()
    except Exception as e:
        print(e)
    else:
        print(f"Time Take: {toc-tic} seconds")
        print("--------------------------------------------")
        

move_files_to_output_folder()
        