import openpyxl
from openpyxl import load_workbook
import easygui
import sys
import re

def prompt_continue(msg):
    title = 'Please confirm'
    if easygui.ccbox(msg, title):
        pass
    else:
        print('User cancelled action')
        sys.exit(0)


def select_workbook():
        path = easygui.fileopenbox(msg='Please select a ".xlsx" or "csv" file', title='Load File')
        try:
            wb = load_workbook(path)
        except:
            print('Selected file not ".xlsx" file type')
            sys.exit(0)
        else:
            return wb


def load_point_codes_list():
    prompt_continue('''    First, select your code file.
    
    This file contains the point codes to be compared against (in column A)''')
    wb = select_workbook()
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
            invalid_codes_present = True
        codes.append(code.value)

    if (invalid_codes_present == True):
        if easygui.ccbox(msg='Invalid codes present, continue?', title='Please Confirm'):
            pass
        else:
            sys.exit(0)

    return codes


def load_survey_points():
    prompt_continue('''    Please select your survey file
                    
    This file can be raw from the surveyor, but must be in .xlsx format''')
    wb = select_workbook()
    ws = wb.active

    #parse survey points into lists [(pt_num, pt_desc)...] 
    def parse_points():
        colA = ws['A']
        colE = ws['E']
        errors_list = []
        point_list = []
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
                    point_list.append((pt_num, desc))
                except:
                    errors_list.append((pt_num, desc))
            else:
                errors_list.append((pt_num, None))

        return {'error_list': errors_list, 'point_list': point_list, 'parsed_desc_point_list': parsed_desc_point_list}
    
    return parse_points()
    

def check_descriptions_against_codes(codes, points):
    unknown_points_list = []
    for point in points['parsed_desc_point_list']:
        for desc in point[1]:
            if desc not in codes:
                unknown_points_list.append(point[0])
                continue
    return unknown_points_list