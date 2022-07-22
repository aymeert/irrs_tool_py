from openpyxl import load_workbook
from openpyxl.styles import Font

path_mac_ar_in = "/Users/aymeerodriguez/Documents/GitHub/irrs_tool_py/example.xlsx"
path_mac_ar_out = "/Users/aymeerodriguez/Documents/GitHub/irrs_tool_py/example_changed.xlsx"
path_mac_jm = "/Users/javier/Documents/GitHub/irrs_tool_py/example.xlsx"
path_mac_jm2 = "/Users/javier/Documents/GitHub/irrs_tool_py/example_changed.xlsx"
path = "C:\\Users\\aymee.rodriguez\\OneDrive - Exactech, Inc\\Projects\\irrs_tool_py\\example.xlsx"
path2 = "C:\\Users\\aymee.rodriguez\\OneDrive - Exactech, Inc\\Projects\\irrs_tool_py\\example_changed.xlsx"
path_full_irss = "C:\\Users\\aymee.rodriguez\\OneDrive - Exactech, Inc\\Projects\\irrs_tool_py\\QC322-110-00 Rev A 2022-07-15-15-22-38.xlsx"
path_full_irss_mac = "/Users/javier/Documents/GitHub/irrs_tool_py/QC322-110-00 Rev A 2022-07-15-15-22-38.xlsx"
path_full_irss_ar = "C:\\Users\\aymee.rodriguez\\OneDrive - Exactech, Inc\\Projects\\irrs_tool_py\\QC322-110-00 Rev A changed 2022-07-15-15-22-38.xlsx"

path_to_translation_table_mac = "/Users/javier/Documents/GitHub/irrs_tool_py/translation_table.xlsx"
path_to_translation_table_win = "C:\\Users\\aymee.rodriguez\\OneDrive - Exactech, Inc\\Projects\\irrs_tool_py\\translation_table.xlsx"

def open_workbook(path_to_workbook):
    """Opens the IRRS to be translated"""
    workbook = load_workbook(path_to_workbook)
    worksheet = workbook["Final Or Supplier Manufactured"]
    return workbook, worksheet


def find_bp_specification(worksheet):
    """Finds the column header to be translated"""
    start_cell = "BP Specification"
    for col in range(worksheet.min_column, worksheet.max_column):
        for row in range(worksheet.min_row, worksheet.max_row):
            if worksheet.cell(row,col).value == start_cell:
                start_row, start_col = row, col
                break
    return start_row, start_col


def iterate_through_column(worksheet):
    """Accesses cell by cell in the column to be translated"""
    start_row, start_col = find_bp_specification(worksheet)
    for row in range(start_row + 1, worksheet.max_row):
        cell = worksheet.cell(row,start_col)
        if cell.value is None:
            break
        cell.font = Font(name= 'Y14.5-2009', size=13)
        translated_symbols = translate_by_cell_type(cell)
        cell.value = translated_symbols
    return worksheet


def frame_simple_cell(cell):
    """Adds framing to cells that require a single GD&T frame"""
    translated_symbols = translate_gdt_symbols(cell)
    translated_symbols = translated_symbols.replace("|", "{", 1)
    translated_symbols = translated_symbols[::-1].replace("|", "}", 1)
    translated_symbols = translated_symbols[::-1]
    before_braket = translated_symbols.split('{')[0]
    after_braket = translated_symbols.split('{')[1]
    #need to brake this off as it's own function
    alphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz."
    complete_translation = ""
    for letter in after_braket:
        if letter in alphabet:
            letter = letter + "_"
        complete_translation = complete_translation + letter
    complete_translation = before_braket + '{' + complete_translation
    return complete_translation


def add_frames_to_characters():
    """Adds frames to individual alpha numeric characters and special GD&T symbols"""
    


def get_list_of_gdt_symbols():
    """Gets a list of all the GD&T symbols from the translation table"""
    translation_table = read_translation_table(path_to_translation_table_win) # change the path depending on the OS


def read_translation_table(path_to_translation_table):
    """Reads the translation table containing the GD&T symbols"""
    workbook = load_workbook(path_to_translation_table)
    translation_table = workbook["Translations"]
    return translation_table


def translate_gdt_symbols(cell):
    """Translates the GD&T symbols from the translation table"""
    translated_symbols = str(cell.value)
    translation_table = read_translation_table(path_to_translation_table_win) # change the path depending on the OS
    for row in range(translation_table.min_row + 1, translation_table.max_row):
        correct_symbol = str(translation_table.cell(row,2).value)
        incorrect_symbol = translation_table.cell(row,3).value #not reading it as a string initially
        if incorrect_symbol:
            translated_symbols = translated_symbols.replace(str(incorrect_symbol), correct_symbol + "~", 1)
            # need to remove adding a ~, this should be reserved for the simple frame function
    return translated_symbols


def translate_by_cell_type(cell):
    """Identifies the kind of cell to use the appropriate translation method"""
    cell_content = cell.value
    if is_simple_frame(cell_content):
        translated_cell = frame_simple_cell(cell)
    else: translated_cell = cell_content
    return translated_cell


def is_simple_frame(cell_content):
    """Identifies if a cell's content is a simple GD&T frame"""
    if cell_content.count("|") >= 2:
        return True
    else: return False

workbook, worksheet  = open_workbook(path_full_irss)
translated_worksheet = iterate_through_column(worksheet)
workbook.save(path2)

# TODO:
"""
    [X] add a function to read an excel table with the codes
    translation
    [X] modify the processIRRS function to replace the symbols based
    on the translation table
    [X] add a function to read each IRRS to be translated
    [X] add a function to export each translated IRRS
    [] add a function to create the exported file with the same name as orignal and in the same folder
    [] add a function to get the active directory and use the translation table in that directory
    [] add 5 cases:
        [X] simple frame
        [] double frame
        [] not framed but translated
        [] Ra
        [X] nothing needs to happen
    [] create a graphical user interface
    [] create an executable program to distribute
"""