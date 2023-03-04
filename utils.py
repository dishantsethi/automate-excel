from config import INPUT_DIR, OUTPUT_DIR, TEXT_DATA_FOR_ROW_TWO, MONTH_YEAR_CELL
import os
from openpyxl.styles import Font, PatternFill
from colors import *

def listdir_nohidden(path):
    for f in os.listdir(path):
        if not f.startswith('.'):
            yield f

def get_sheets_in_dir(dir):
    return listdir_nohidden(dir)

def get_visible_sheet_list(wb):
    sheet_list = []
    for sheet in wb.sheetnames:
        if wb[sheet].sheet_state == 'visible':
            sheet_list.append(sheet)
    return sheet_list

def update_font(ws, size):
    try:
        for rows in ws.iter_cols():
            for index, cell in enumerate(rows):
                name = cell.font.name
                charset = cell.font.charset
                family = cell.font.family
                b = cell.font.b
                i = cell.font.i
                strike = cell.font.strike
                outline = cell.font.outline
                shadow = cell.font.shadow
                condense = cell.font.condense
                color = cell.font.color
                cell.font = Font(name=name, charset=charset, family=family, b=b, i=i, strike=strike, outline=outline, shadow=shadow, condense=condense, size=size, color=color)

                patternType = cell.fill.patternType
                cell.fill = PatternFill(patternType=patternType, fgColor="FFFFFFFF")

                if index in [0, 1, 2]:
                    cell.font = Font(name=name, charset=charset, family=family, b=True, i=i, strike=strike, outline=outline, shadow=shadow, condense=condense, size=size, color=color)
    except Exception as e:
        print_bold_red(f"Unable to update font: {e}")

def update_summary_row(ws, a2, bankname):
    try:
        ws.insert_rows(2)
        ws["A2"].value = f"{TEXT_DATA_FOR_ROW_TWO} {a2} {bankname}"
    except Exception as e:
        print_bold_red(f"Unable to insert row: {e}")

def move_files_to_output_folder(files):
    for file in files:
        source = os.path.join(INPUT_DIR, file)
        des = os.path.join(OUTPUT_DIR, file)
        os.rename(source, des)
        
def get_year_and_month_for_a2(wb):
    a2_data = wb["Summary"][MONTH_YEAR_CELL].value
    if a2_data:
        data = "{:%B %Y}".format(a2_data)
        print_bold_header(f"Year and Month: {data}")
        return data    
    print_bold_red(f"Year and Month Missing in cell {MONTH_YEAR_CELL}")
    return a2_data

def get_bankname(file):
    l = file.split(" ")
    bankname = f"{l[2]} {l[4]} {l[5]}"
    print_bold_header(f"Bank name to be appended in inserted row {bankname}")
    return bankname