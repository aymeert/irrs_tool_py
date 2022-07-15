from openpyxl.styles import Font
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook

path_mac_jm = "/Users/javier/Documents/GitHub/irrs_tool_py/example.xlsx"
path_mac_jm2 = "/Users/javier/Documents/GitHub/irrs_tool_py/example_changed.xlsx"
path = "C:\\Users\\aymee.rodriguez\\OneDrive - Exactech, Inc\\Projects\\irrs_tool_py\\example.xlsx"
path2 = "C:\\Users\\aymee.rodriguez\\OneDrive - Exactech, Inc\\Projects\\irrs_tool_py\\example_changed.xlsx"

wb = load_workbook(path_mac_jm)
ws = wb.active
cell = ws.cell(row = 1, column = 1)
wrong_symbols = str(cell.value)
right_symbols = wrong_symbols.replace("|", "{", 1)
right_symbols = right_symbols[::-1].replace("|", "}", 1)
right_symbols = right_symbols[::-1]
right_symbols = right_symbols.replace("⌖", "¿~", 1)
right_symbols = right_symbols.replace("Ⓜ", "Ì~", 1)
print(right_symbols)
ws['A1'] = right_symbols
ws['A1'].font = Font(name= 'Y14.5-2009')

wb.save(path_mac_jm2)
# TODO:
    # [] add a function to read an excel table with the codes
    #   translation
    # [] modify the processIRRS function to replace the symbols based
    #   on the translation table
    # [] add a function to read each IRRS to be translated
    # [] add a function to export each translated IRRS