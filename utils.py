from config import INPUT_DIR, OUTPUT_DIR
import os
from openpyxl.styles import Font, PatternFill

def listdir_nohidden(path):
    for f in os.listdir(path):
        if not f.startswith('.'):
            yield f

def get_sheets_from_input_dir():
    return listdir_nohidden(INPUT_DIR)

def get_visible_sheet_list(wb):
    sheet_list = []
    for sheet in wb.sheetnames:
        if wb[sheet].sheet_state == 'visible':
            sheet_list.append(sheet)
    return sheet_list

def update_font(cell, index, size, sheet):
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

def move_files_to_output_folder():
    for file in os.listdir(INPUT_DIR):
        if file.endswith(".xlsx"):
            source = os.path.join(INPUT_DIR, file)
            des = os.path.join(OUTPUT_DIR, file)
            os.rename(source, des)