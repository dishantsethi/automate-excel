from config import INPUT_DIR, SIZE, TEXT_DATA_FOR_ROW_TWO
from utils import get_visible_sheet_list, update_font

def page_setup_for_each_sheet(wb, wb_name, a2, bankname):
    sheet_list = get_visible_sheet_list(wb)
    for sheet in sheet_list:
        ws = wb[sheet]
        
        print(f"Insert row 2 in sheet {sheet}")
        ws.insert_rows(2)
        ws["A2"].value = f"{TEXT_DATA_FOR_ROW_TWO} {a2} {bankname}"
        
        print(f"Updating font of sheet {sheet} for workbook {wb_name}")
        size = SIZE[sheet] if sheet in SIZE else SIZE["default"]
        for rows in ws.iter_cols():
            for index, cell in enumerate(rows):
                update_font(cell, index, size, sheet)
    wb.save(f"{INPUT_DIR}/{wb_name}")
    