from config import INPUT_DIR, FONT_SIZE
from utils import get_visible_sheet_list, update_font, update_summary_row

def page_setup_for_each_sheet(wb, wb_name, a2, bankname):
    sheet_list = get_visible_sheet_list(wb)
    for sheet in sheet_list:
        ws = wb[sheet]
        size = FONT_SIZE[sheet] if sheet in FONT_SIZE else FONT_SIZE["default"]
        
        print(f"Inserting row 2 in sheet {sheet}")
        update_summary_row(ws, a2, bankname)
        
        print(f"Updating font of sheet {sheet}")        
        update_font(ws, size)

    wb.save(f"{INPUT_DIR}/{wb_name}")
