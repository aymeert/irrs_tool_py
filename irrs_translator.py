from openpyxl.styles import Font
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook


path_mac_ar_in = "/Users/aymeerodriguez/Documents/GitHub/irrs_tool_py/example.xlsx"
path_mac_ar_out = "/Users/aymeerodriguez/Documents/GitHub/irrs_tool_py/example_changed.xlsx"

path_mac_jm = "/Users/javier/Documents/GitHub/irrs_tool_py/example.xlsx"
path_mac_jm2 = "/Users/javier/Documents/GitHub/irrs_tool_py/example_changed.xlsx"
path = "C:\\Users\\aymee.rodriguez\\OneDrive - Exactech, Inc\\Projects\\irrs_tool_py\\example.xlsx"
path2 = "C:\\Users\\aymee.rodriguez\\OneDrive - Exactech, Inc\\Projects\\irrs_tool_py\\example_changed.xlsx"

path_full_irss = "C:\\Users\\aymee.rodriguez\\OneDrive - Exactech, Inc\\Projects\\irrs_tool_py\\QC322-110-00 Rev A 2022-07-15-15-22-38.xlsx"
path_full_irss_mac = "/Users/javier/Documents/GitHub/irrs_tool_py/QC322-110-00 Rev A 2022-07-15-15-22-38.xlsx"
path_full_irss_ar = "C:\\Users\\aymee.rodriguez\\OneDrive - Exactech, Inc\\Projects\\irrs_tool_py\\QC322-110-00 Rev A changed 2022-07-15-15-22-38.xlsx"

def open_workbook(path_to_workbook):                                     # opening excel 
    workbook = load_workbook(path_to_workbook)
    worksheet = workbook["Final Or Supplier Manufactured"]
    return workbook, worksheet

def find_bp_specification(worksheet):                                    # finding the first cell to translate
    start_cell = "BP Specification"
    for col in range(worksheet.min_column, worksheet.max_column):
        for row in range(worksheet.min_row, worksheet.max_row):
            if worksheet.cell(row,col).value == start_cell:
                start_row, start_col = row, col
                break
    return start_row, start_col

def iterate_through_column(worksheet):                                   # function to keep translating the following cells 
    start_row, start_col = find_bp_specification(worksheet)
    for row in range(start_row + 1, worksheet.max_row):
        cell = worksheet.cell(row,start_col)
        if cell.value is None:
            break
        cell.font = Font(name= 'Y14.5-2009')
        translated_symbols = translate_by_cell_type(cell)
        cell.value = translated_symbols
    return worksheet

def frame_simple_cell(cell):                                            # function to create the box for the perfect dimensions
    wrong_symbols = str(cell.value)
    wrong_symbols = str(cell.value)
    right_symbols = wrong_symbols.replace("|", "{", 1)
    right_symbols = right_symbols[::-1].replace("|", "}", 1)
    right_symbols = right_symbols[::-1]
    right_symbols = right_symbols.replace("⌖", "¿~", 1)
    right_symbols = right_symbols.replace("Ⓜ", "Ì~", 1)
    cell.font = Font(name= 'Y14.5-2009')
    first_half = right_symbols.split('{')[0]
    second_half = right_symbols.split('{')[1]
    alphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz."
    final_symbols = ""
    for letter in second_half:
        if letter in alphabet:
            letter = letter + "`"
        final_symbols = final_symbols + letter
    final_symbols = first_half+'{'+final_symbols
    return final_symbols

def translate_by_cell_type(cell):                                          # Main funtion for identification of the type of cell (so far only two cases: simple cell and no change)  
    cell_content = cell.value
    if is_simple_frame(cell_content):
        translated_cell = frame_simple_cell(cell)
    else: translated_cell = cell_content
    return translated_cell

def is_simple_frame(cell_content):
    if cell_content.count("|") >= 2:
        return True
    else: return False

workbook, worksheet  = open_workbook(path_full_irss)                   # exporting into another excel file
translated_worksheet = iterate_through_column(worksheet)
workbook.save(path2)

# TODO:
    # [] add a function to read an excel table with the codes
    #   translation
    # [] modify the processIRRS function to replace the symbols based
    #   on the translation table
    # [X] add a function to read each IRRS to be translated
    # [X] add a function to export each translated IRRS
    # [] create a user interface