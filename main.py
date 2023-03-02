from openpyxl import load_workbook
import os
from config import INPUT_DIR
from utils import get_sheets_from_input_dir, move_files_to_output_folder
from functions import page_setup_for_each_sheet
import time

count = 0
for file in get_sheets_from_input_dir():
    filename = os.path.join(INPUT_DIR, file)
    try:
        wb = load_workbook(filename = filename)
        print("--------------------------------------------")
        print(f"File name {file}")
        tic = time.time()
        a2 = "{:%B %d %Y}".format(wb["Summary"]["C2"].value)
        l = file.split(" ")
        bankname = f"{l[2]} {l[4]} {l[5]}"
        page_setup_for_each_sheet(wb, file, a2, bankname)
        toc = time.time()
    except Exception as e:
        print(e)
    else:
        print(f"Time Take: {toc-tic} seconds")
        print("--------------------------------------------")
        count = count + 1

if len(list(get_sheets_from_input_dir())) == count:
    move_files_to_output_folder()
    