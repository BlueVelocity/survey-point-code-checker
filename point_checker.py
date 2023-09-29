import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
import easygui
import sys
import re

def unidentified_error_handler(func):
    def wrapper(*args):
        try:
            return func(*args)
        except:
            sys.exit(f'ExecutionError: Action stopped at "{func.__name__}"')
    return wrapper

@unidentified_error_handler
def prompt_continue(msg):
    title = 'Please confirm'
    if easygui.ccbox(msg, title):
        pass
    else:
        print('User action cancelled')
        sys.exit(0)

@unidentified_error_handler
def prompt_string_input(msg):
    title = 'Please enter...'
    d_text = 'output'
    check_file_name_validity_pattern = r"^[A-Za-z0-9]+([-_][A-Za-z0-9]+)*(\.[A-Za-z0-9]+)?$"
    is_valid = False
    while(is_valid == False):
        string = easygui.enterbox(msg, title, d_text)
        if (re.match(check_file_name_validity_pattern, string) != None):
            is_valid = True
        elif (string == None):
            print('User entered nothing')
            sys.exit(0)
    return string        

@unidentified_error_handler
def select_workbook():
        path = easygui.fileopenbox(msg='Please select a ".xlsx" or "csv" file', title='Load File')
        try:
            wb = load_workbook(path)
        except:
            print('Selected file not ".xlsx" file type')
            sys.exit(0)
        else:
            return wb

@unidentified_error_handler
def load_point_codes_list():
    prompt_continue('''    First, select your code file.
    
    This file contains the point codes to be compared against (in column A)''')
    wb = select_workbook()
    ws = wb.active

    def valid_code(code):
        pattern = r'[0-9]'
        try:
            if re.search(pattern, code):
                return False
            else:
                return True
        except:
            print('Issue with code point format. Is this a code file?')
            sys.exit(0)
        
    codes = []
    invalid_codes_present = False
    column_to_check = 'A'

    for code in ws[column_to_check]:
        if(code.value == None):
            continue
        if (valid_code(code.value) == False):
            print(f'WARNING Invalid Code: {code}')
            invalid_codes_present = True
        codes.append(code.value)

    if (invalid_codes_present == True):
        if easygui.ccbox(msg='Invalid codes present, continue?', title='Please Confirm'):
            pass
        else:
            sys.exit(0)

    return codes

@unidentified_error_handler
def load_survey_points():
    prompt_continue('''    Please select your survey file
                    
    This file can be raw from the surveyor, but must be in .xlsx format''')
    wb = select_workbook()
    ws = wb.active

    #parse survey points into lists [(pt_num, pt_desc)...] 
    def parse_points():
        colA = ws['A']
        colE = ws['E']
        error_list = []
        point_list = {}
        parsed_desc_point_list = []
        for num in range(len(colA)):
            pt_num = colA[num].value
            desc = colE[num].value

            def parse_description(desc):
                remove_numbers_and_decimals_pattern = r'-?\d+(\.\d+)?'
                remove_multiplication_symbol = r'\s[xX]\s'
                combined_pattern = r'|'.join((remove_multiplication_symbol, remove_numbers_and_decimals_pattern))
                sanitized_desc = re.sub(combined_pattern, '', desc)
                parsed_desc = sanitized_desc.split()
                return parsed_desc
            
            if (desc != None):
                try:
                    parsed_desc = parse_description(desc)
                    parsed_desc_point_list.append((pt_num, parsed_desc))
                    point_list[f'{pt_num}'] = desc
                except:
                    error_list.append((pt_num, desc))
            else:
                error_list.append((pt_num, None))

        return {'error_list': error_list, 'point_list': point_list, 'parsed_desc_point_list': parsed_desc_point_list}
    
    return parse_points()
    
@unidentified_error_handler
def check_descriptions_against_codes(codes, points):
    unknown_points_list = []
    for point in points['parsed_desc_point_list']:
        for desc in point[1]:
            if desc not in codes:
                unknown_points_list.append(point[0])
                continue
    return unknown_points_list

@unidentified_error_handler
def output_points(points, unknown_points):
    name = easygui.filesavebox()

    wb = Workbook()
    ws = wb.active

    for index, pt_num in enumerate(unknown_points, 1):
        point_list = points['point_list']
        pt_desc = point_list[f'{pt_num}']
        ws[f'A{index}'] = pt_num
        ws[f'E{index}'] = pt_desc
    
    wb.save(f'{name}.xlsx')
    
point_codes = load_point_codes_list()
survey_points = load_survey_points()
unknown_points = check_descriptions_against_codes(point_codes, survey_points)

output_points(survey_points, unknown_points)