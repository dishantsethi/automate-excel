from config import PDF_DIR, OUTPUT_DIR, summary, load_book_movement, prepayments_and_reschedulement, collections_and_overdues
from utils import get_sheets_in_dir, get_visible_sheet_list, get_sheet_row_count
from openpyxl import load_workbook
from openpyxl.worksheet.page import PageMargins, PrintPageSetup
import os
from openpyxl.styles import Border
from colors import print_bold_header, print_bold_blue, print_bold_warning, print_bold_green
import win32com.client
from pywintypes import com_error
import time
import shutil

def convert_to_pdf():

    generate_temp_excel_for_pdf()
    update_temp_excel_and_convert_to_pdf()
    

def generate_temp_excel_for_pdf():
    for file in get_sheets_in_dir(OUTPUT_DIR):
        source = os.path.join(OUTPUT_DIR, file)
        des = os.path.join(PDF_DIR, f"tmp{file}")
        shutil.copy(source, des)

def update_temp_excel_and_convert_to_pdf():
    for file in get_sheets_in_dir(PDF_DIR):
        print_bold_blue("------------------------------------------------------------")
        tic = time.time()
        if file.endswith(".pdf"):
            continue
        filename = os.path.join(PDF_DIR, file)
        wb = load_workbook(filename=filename)
        sheet_list = get_visible_sheet_list(wb)
        for sheet in sheet_list:
            ws = wb[sheet]
            if sheet == summary:
                ws.page_margins = PageMargins(left=0.50, right=0.50, top=0.50, bottom=1.50, header=0.3, footer=0.3)
                ws.page_setup = PrintPageSetup(orientation=ws.ORIENTATION_PORTRAIT, paperSize=ws.PAPERSIZE_A4, scale=65)
            elif sheet == load_book_movement:
                ws.page_margins = PageMargins(left=0.50, right=0.50, top=0.50, bottom=1.50, header=0.3, footer=0.3)
                ws.page_setup = PrintPageSetup(orientation=ws.ORIENTATION_LANDSCAPE, paperSize=ws.PAPERSIZE_A4, scale=65)
                update_border(ws)
            elif sheet == prepayments_and_reschedulement:
                ws.page_margins = PageMargins(left=0.50, right=0.50, top=0.50, bottom=1.50, header=0.3, footer=0.3)
                ws.page_setup = PrintPageSetup(orientation=ws.ORIENTATION_LANDSCAPE, paperSize=ws.PAPERSIZE_A4, scale=65)
            elif sheet == collections_and_overdues:
                ws.page_margins = PageMargins(left=0.50, right=0.50, top=0.50, bottom=1.50, header=0.3, footer=0.3)
                ws.page_setup = PrintPageSetup(orientation=ws.ORIENTATION_LANDSCAPE, paperSize=ws.PAPERSIZE_A4, scale=60)
                update_border(ws)
            else:
                ws.page_margins = PageMargins(left=0.50, right=0.50, top=0.50, bottom=1.50, header=0.3, footer=0.3)
                ws.page_setup = PrintPageSetup(orientation=ws.ORIENTATION_PORTRAIT, paperSize=ws.PAPERSIZE_A4, scale=65)
        wb.save(filename)
        print_bold_header(f"Page Setup Done for file {file}")
        print_bold_header(f"Borders removed for file {file}")
        excel_to_pdf(file)
        toc = time.time()
        print_bold_green(f"Time Take: {toc-tic} seconds")
        print_bold_blue("------------------------------------------------------------")


def update_border(ws):
    max_row = get_sheet_row_count(ws)
    for rows in ws.iter_rows(min_row=1, min_col=1, max_row=max_row ,max_col=ws.max_column):
        for cell in rows:
            cell.border = Border(left=None, right=None, bottom=None, top=None, outline=None)
        
def excel_to_pdf(excel_file):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.interactive = False
    excel.visible = False
    pdf_file = excel_file[3:len(excel_file)-4] + "pdf"
    pdf_file_path = os.path.join(PDF_DIR, pdf_file)
    excel_file_path = os.path.join(PDF_DIR, excel_file)
    try:
        wb = excel.Workbooks.Open(excel_file_path)
        ws_index_list = [1,2,3,4]
        wb.WorkSheets(ws_index_list).Select()
        wb.ActiveSheet.ExportAsFixedFormat(0, pdf_file_path)
    except com_error as e:
        print(f'Failed. {e}')
    else:
        print_bold_warning(f"Converted {excel_file} to {pdf_file}")
    finally:
        wb.Close()
        excel.Quit()
    
