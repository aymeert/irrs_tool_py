from openpyxl import load_workbook
from openpyxl.styles import Font
import openpyxl.cell._writer
from pathlib import Path

path_to_translation_table_mac = "/Users/javier/Documents/GitHub/irrs_tool_py/translation_table.xlsx"
path_to_translation_table = Path("J:\\Public\\Employee\\AYMEE.RODRIGUEZ\\IRRS translator program\\translation_table.xlsx")

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
    for row in range(start_row + 1, worksheet.max_row + 1):
        cell = worksheet.cell(row,start_col)
        if cell.value is None:
            break
        cell.font = Font(name= 'Y14.5-2009', size=13)
        translated_symbols = translate_by_cell_type(cell)
        cell.value = translated_symbols
    return worksheet


def translate_by_cell_type(cell):
    """Identifies the kind of cell to use the appropriate translation method"""
    cell_content = cell.value
    if is_simple_frame(cell_content):
        translated_cell = frame_simple_cell(cell)
    else: translated_cell = translate_gdt_symbols(cell) # this was changed from cell_content
    return translated_cell


def is_simple_frame(cell_content):
    """Identifies if a cell's content is a simple GD&T frame"""
    if cell_content.count("|") >= 2:
        return True
    else: return False


def frame_simple_cell(cell):
    """Adds framing to cells that require a single GD&T frame"""
    translated_symbols = translate_gdt_symbols(cell)
    translated_symbols = translated_symbols.replace("|", "{", 1)
    translated_symbols = translated_symbols[::-1].replace("|", "}", 1)
    translated_symbols = translated_symbols[::-1]
    before_first_braket = translated_symbols.split('{')[0]
    after_first_braket = translated_symbols.split('{')[1]
    before_second_braket = after_first_braket.split('}')[0]
    after_second_braket = after_first_braket.split('}')[1]
    framed_characters = add_frames_to_characters(before_second_braket)
    complete_translation = before_first_braket + '{' + framed_characters + '}' + after_second_braket
    return complete_translation


def translate_gdt_symbols(cell):
    """Translates the GD&T symbols from the translation table"""
    translated_symbols = str(cell.value)
    translation_table = read_translation_table(path_to_translation_table) # change the path depending on the OS
    for row in range(translation_table.min_row + 1, translation_table.max_row):
        correct_symbol = str(translation_table.cell(row,2).value)
        incorrect_symbol = translation_table.cell(row,3).value #not reading it as a string initially
        if incorrect_symbol:
            translated_symbols = translated_symbols.replace(str(incorrect_symbol), correct_symbol)
    return translated_symbols


def read_translation_table(path_to_translation_table):
    """Reads the translation table containing the GD&T symbols"""
    workbook = load_workbook(path_to_translation_table)
    translation_table = workbook["Translations"]
    return translation_table


def add_frames_to_characters(characters_to_frame):
    """Adds frames to individual alpha numeric characters and special GD&T symbols"""
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    numbers = "0123456789"
    punctuation = "."
    gdt_symbols = get_list_of_gdt_symbols(path_to_translation_table)
    framed_characters = ""
    for character in characters_to_frame:
        if character in alphabet:
            character = character + "_"
        elif character in numbers:
            character = character + "`"
        elif character in punctuation:
            character = character + "\\"
        elif character in gdt_symbols:
            character = character + "~"
        framed_characters = framed_characters + character
    return framed_characters


def get_list_of_gdt_symbols(path_to_translation_table):
    """Gets a list of all the GD&T symbols from the translation table"""
    translation_table = read_translation_table(path_to_translation_table) # change the path depending on the OS
    gdt_symbols = ""
    for row in range(translation_table.min_row + 1, translation_table.max_row):
        incorrect_symbol = translation_table.cell(row,3).value #not reading it as a string initially
        if incorrect_symbol:
            gdt_symbols = gdt_symbols + str(translation_table.cell(row,2).value)
    return gdt_symbols


def generate_irrs_output_path(path_to_irrs_for_translation):
    """Generate the output path for the translated IRRS with the same name"""
    path_to_translated_irrs = path_to_irrs_for_translation.replace(".xlsx", " [Translated].xlsx")
    path_to_translated_irrs = Path(path_to_translated_irrs) # !!! Need to add [1:-1] at the end of the variable !!!
    return path_to_translated_irrs

path_to_irrs_for_translation = "C:\\Users\\aymee.rodriguez\\OneDrive - Exactech, Inc\\Projects\\irrs_tool_py\\QC321-150-46 Rev A 2022-07-08-10-10-16.xlsx"

path_to_translated_irrs = generate_irrs_output_path(path_to_irrs_for_translation)
workbook, worksheet  = open_workbook(Path(path_to_irrs_for_translation))
translated_worksheet = iterate_through_column(worksheet)
workbook.save(path_to_translated_irrs)

# print(r"""
#  _______   __  ___  _____ _____ _____ _____  _   _                                        
# |  ___\ \ / / / _ \/  __ \_   _|  ___/  __ \| | | |                                       
# | |__  \ V / / /_\ \ /  \/ | | | |__ | /  \/| |_| |                                       
# |  __| /   \ |  _  | |     | | |  __|| |    |  _  |                                       
# | |___/ /^\ \| | | | \__/\ | | | |___| \__/\| | | |                                       
# \____/\/   \/\_| |_/\____/ \_/ \____/ \____/\_| |_/                                       
                                                                                          
                                                                                          
#  _________________  _____   ___________  ___   _   _  _____ _       ___ _____ ___________ 
# |_   _| ___ \ ___ \/  ___| |_   _| ___ \/ _ \ | \ | |/  ___| |     / _ \_   _|  _  | ___ \
#   | | | |_/ / |_/ /\ `--.    | | | |_/ / /_\ \|  \| |\ `--.| |    / /_\ \| | | | | | |_/ /
#   | | |    /|    /  `--. \   | | |    /|  _  || . ` | `--. \ |    |  _  || | | | | |    / 
#  _| |_| |\ \| |\ \ /\__/ /   | | | |\ \| | | || |\  |/\__/ / |____| | | || | \ \_/ / |\ \ 
#  \___/\_| \_\_| \_|\____/    \_/ \_| \_\_| |_/\_| \_/\____/\_____/\_| |_/\_/  \___/\_| \_|
                                                                                                                                                                                    
# """)

# while True:
#     path_to_irrs_for_translation = input("Drag the IRRS to be translated in here, and press enter to translate: ")
#     if path_to_irrs_for_translation:
#         path_check_if_valid_irrs = Path(path_to_irrs_for_translation[1:-1])
#         if path_check_if_valid_irrs.is_file():
#             path_to_translated_irrs = generate_irrs_output_path(path_to_irrs_for_translation)
#             workbook, worksheet  = open_workbook(Path(path_to_irrs_for_translation[1:-1]))
#             translated_worksheet = iterate_through_column(worksheet)
#             workbook.save(path_to_translated_irrs)
#             print("Translation successful! New file saved in the same location with [Translated] appended")
#         else:
#             print("Path provided is not valid, try again")
#     else:
#         print("Input is not valid, try again")

# TODO:
"""
    [X] add a function to read an excel table with the codes
    translation
    [X] modify the processIRRS function to replace the symbols based
    on the translation table
    [X] add a function to read each IRRS to be translated
    [X] add a function to export each translated IRRS
    [] replace all string concatenation with ''.join()
    [X] add a function to create the exported file with the same name as orignal and in the same folder
    [X] add a function to get the active directory and use the translation table in that directory
    [] add 5 cases:
        [X] simple frame
            [X] simple frame with text at the end
        [] double frame
        [X] not framed but translated
        [] Ra
        [X] nothing needs to happen
    [x] create a graphical (or terminal) user interface
    [X] create an executable program to distribute
"""