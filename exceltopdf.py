from config import PDF_DIR, OUTPUT_DIR
from utils import get_sheets_in_dir, get_visible_sheet_list
from openpyxl import load_workbook
from openpyxl.worksheet.page import PageMargins, PrintPageSetup
import os
from openpyxl.styles import Border
from colors import print_bold_header


def convert_to_pdf():
    generate_temp_excel_for_pdf()
    update_temp_excel_and_convert_to_pdf()
    

def generate_temp_excel_for_pdf():
    for file in get_sheets_in_dir(OUTPUT_DIR):
        filename = os.path.join(OUTPUT_DIR, file)
        wb = load_workbook(filename=filename)
        wb.save(f"{PDF_DIR}/{file}")

def update_temp_excel_and_convert_to_pdf():
    for file in get_sheets_in_dir(PDF_DIR):
        filename = os.path.join(PDF_DIR, file)
        wb = load_workbook(filename=filename)
        sheet_list = get_visible_sheet_list(wb)
        for sheet in sheet_list:
            ws = wb[sheet]
            if sheet == "Summary":
                ws.page_margins = PageMargins(left=0.50, right=0.50, top=0.50, bottom=1.50, header=0.3, footer=0.3)
                ws.page_setup = PrintPageSetup(orientation=ws.ORIENTATION_PORTRAIT, paperSize=ws.PAPERSIZE_A4, scale=65)
            elif sheet == "Loan Book Movement":
                ws.page_margins = PageMargins(left=0.50, right=0.50, top=0.50, bottom=1.50, header=0.3, footer=0.3)
                ws.page_setup = PrintPageSetup(orientation=ws.ORIENTATION_LANDSCAPE, paperSize=ws.PAPERSIZE_A4, scale=65)
                update_border(ws)
            elif sheet == "Prepayments & Reschedulement":
                ws.page_margins = PageMargins(left=0.50, right=0.50, top=0.50, bottom=1.50, header=0.3, footer=0.3)
                ws.page_setup = PrintPageSetup(orientation=ws.ORIENTATION_LANDSCAPE, paperSize=ws.PAPERSIZE_A4, scale=65)
            elif sheet == "Collections & Overdues":
                ws.page_margins = PageMargins(left=0.50, right=0.50, top=0.50, bottom=1.50, header=0.3, footer=0.3)
                ws.page_setup = PrintPageSetup(orientation=ws.ORIENTATION_LANDSCAPE, paperSize=ws.PAPERSIZE_A4, scale=60)
                update_border(ws)
            else:
                ws.page_margins = PageMargins(left=0.50, right=0.50, top=0.50, bottom=1.50, header=0.3, footer=0.3)
                ws.page_setup = PrintPageSetup(orientation=ws.ORIENTATION_PORTRAIT, paperSize=ws.PAPERSIZE_A4, scale=65)
        wb.save(filename)
        print_bold_header(f"Page Setup Done for file {file}")
        print_bold_header(f"Borders removed for file {file}")
        excel_to_pdf(filename)

def update_border(ws):
    for rows in ws.iter_rows():
        for cell in rows:
            cell.border = Border(left=None, right=None, bottom=None, top=None, outline=None)
        
def excel_to_pdf(file):
    print(f"Converting {file} to pdf")
