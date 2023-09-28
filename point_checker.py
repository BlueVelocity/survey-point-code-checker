import openpyxl
from openpyxl import workbook
from openpyxl import load_workbook
import easygui
import sys
import re

def select_file(func):
    def wrapper():
        path = easygui.fileopenbox(msg='Please select a ".xlsx" file', title='Load File')
        try:
            wb = load_workbook(path)
        except:
            print('Selected file not ".xlsx" file type')
            sys.exit()
        else:
            func(path)
    return wrapper

@select_file
def load_point_codes_list(path):
    wb = load_workbook(path)
    ws = wb.active
    
    def valid_code(code):
        pattern = r'[0-9]'
        if re.search(pattern, code):
            return False
        else:
            return True

    codes = []
    invalid_codes_present = False
    column_to_check = 'A'
    for code in ws[column_to_check]:
        if(code.value == None):
            continue
        if (valid_code(code.value) == False):
            print(f'WARNING Invalid Code: {code}')
        codes.append(code.value)
    if (invalid_codes_present == True):
        selection = easygui.ynbox(msg='Invalid codes present, continue?')
        if (selection == False):
            sys.exit()
    with open('test.txt', 'w') as text:
        text.write(f'{codes}')