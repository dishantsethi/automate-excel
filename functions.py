from config import INPUT_DIR, FONT_SIZE, summary, loan_book_movement
from utils import get_visible_sheet_list, update_font, insert_row_a2

def page_setup_for_each_sheet(wb, wb_name, a2, bankname):
    sheet_list = get_visible_sheet_list(wb)
    for sheet in sheet_list:
        ws = wb[sheet]
        if sheet in FONT_SIZE:
            size = FONT_SIZE[sheet]
        elif sheet.startswith(summary):
            size = FONT_SIZE[summary]
        elif sheet.startswith("Loan"):
            size = FONT_SIZE[loan_book_movement]
        else:
            size = FONT_SIZE["default"]
 
        print(f"Inserting row 2 in sheet {sheet}")
        insert_row_a2(ws, a2, bankname)
        
        print(f"Updating font of sheet {sheet}")        
        update_font(ws, size)

    wb.save(f"{INPUT_DIR}/{wb_name}")
