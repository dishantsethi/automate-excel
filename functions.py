from config import INPUT_DIR, SIZE, TEXT_DATA_FOR_ROW_TWO
from utils import get_visible_sheet_list, update_font

def insert_second_row_in_summary(wb, wb_name):
    sheet = wb["Summary"]
    print(f"Insert row 2 in sheet {sheet}")
    sheet.insert_rows(2)
    sheet["A2"].value = TEXT_DATA_FOR_ROW_TWO
    wb.save(f"{INPUT_DIR}/{wb_name}")

def convert_sheet_into_pdf(file):
    pass

def page_setup_for_each_sheet(wb, wb_name):
    sheet_list = get_visible_sheet_list(wb)
    for sheet in sheet_list:
        print(f"Updating font of sheet {sheet} for workbook {wb_name}")
        ws = wb[sheet]
        size = SIZE[sheet] if sheet in SIZE else SIZE["default"]
        for rows in ws.iter_cols():
            for index, cell in enumerate(rows):
                update_font(cell, index, size, sheet)
    wb.save(f"{INPUT_DIR}/{wb_name}")
    # print("Converting Sheet into PDF")
    # convert_sheet_into_pdf(f"{INPUT_DIR}/{wb_name}")

