from openpyxl import load_workbook
import os
import time
from config import INPUT_DIR
from utils import get_sheets_in_dir, move_files_to_output_folder, get_bankname, get_year_and_month_for_a2
from functions import page_setup_for_each_sheet
from colors import print_bold_warning, print_bold_blue, print_bold_green
from exceltopdf import convert_to_pdf


def start():
    success = []       
    for file in get_sheets_in_dir(INPUT_DIR):
        filename = os.path.join(INPUT_DIR, file)
        try:
            print_bold_blue("------------------------------------------------------------")
            print_bold_warning(f"File name {file}")
            tic = time.time()
            wb = load_workbook(filename = filename)
            a2_data = get_year_and_month_for_a2(wb)
            bankname = get_bankname(file)
            page_setup_for_each_sheet(wb, file, a2_data, bankname)
            toc = time.time()
        except Exception as e:
            print_bold_warning(e)
        else:
            print_bold_green(f"Time Take: {toc-tic} seconds")
            print_bold_blue("------------------------------------------------------------")
            success.append(file)    
    move_files_to_output_folder(success)
    
if __name__ == "__main__":
    start()
    user_input = input("Do you want to generate PDf for all the files in output folder? (y/n)")
    if user_input in ["y", "Y"]:
        convert_to_pdf()