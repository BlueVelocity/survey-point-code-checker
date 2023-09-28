import openpyxl
from openpyxl import load_workbook
import easygui
import sys
import re

def select_file(func):
    def wrapper():
        path = easygui.fileopenbox(msg='Please select a ".xlsx" or "csv" file', title='Load File')
        try:
            wb = load_workbook(path)
        except:
            print('Selected file not ".xlsx" file type')
            sys.exit()
        else:
            return func(path)
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

    return codes


@select_file  
def load_survey_points(path):
    wb = load_workbook(path)
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

point_codes = load_point_codes_list()
survey_points = load_survey_points()
unkown_points = check_descriptions_against_codes(point_codes, survey_points)