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



def find_bp_specification(worksheet):
    start_cell = "BP Specification"
    for col in range(worksheet.min_column, worksheet.max_column):
        for row in range(worksheet.min_row, worksheet.max_row):
            if worksheet.cell(row,col).value == start_cell:
                start_row, start_col = row, col
                break
    return start_row, start_col

def iterate_through_column(worksheet):
    start_row, start_col = find_bp_specification(worksheet)
    for row in range(start_row, worksheet.max_row):
        print(worksheet.cell(row,start_col).value)


wb = load_workbook(path_full_irss)
ws = wb.active

iterate_through_column(ws)

cell = ws.cell(row = 1, column = 1)
wrong_symbols = str(cell.value)
right_symbols = wrong_symbols.replace("|", "{", 1)
right_symbols = right_symbols[::-1].replace("|", "}", 1)
right_symbols = right_symbols[::-1]
right_symbols = right_symbols.replace("⌖", "¿~", 1)
right_symbols = right_symbols.replace("Ⓜ", "Ì~", 1)
# print(right_symbols)
# ws['A1'] = right_symbols
ws['A1'].font = Font(name= 'Y14.5-2009')

first_half = right_symbols.split('{')[0]
second_half = right_symbols.split('{')[1]

alphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz."
final_symbols = ""
for letter in second_half:
    if letter in alphabet:
        letter = letter + "`"
    final_symbols = final_symbols + letter
final_symbols = first_half+'{'+final_symbols
# print(final_symbols)

ws['A1'] = final_symbols
wb.save(path2)
# TODO:
    # [] add a function to read an excel table with the codes
    #   translation
    # [] modify the processIRRS function to replace the symbols based
    #   on the translation table
    # [] add a function to read each IRRS to be translated
    # [] add a function to export each translated IRRS