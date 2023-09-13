import openpyxl
import re
import easygui

#opens codes file and creates a list of codes for comparison
path_codes = "codes.xlsx"
codes_wb_obj = openpyxl.load_workbook(path_codes)
code_sheet_obj = codes_wb_obj.active
code_rows = code_sheet_obj.max_row

codes = []

for i in range(5, code_rows):
    cell_obj = code_sheet_obj.cell(row = i, column = 1)
    codes.append(cell_obj.value)

#opens survey file and creates a list of points for comparison
path_surv = easygui.fileopenbox()
surv_wb_obj = openpyxl.load_workbook(path_surv)
surv_sheet_obj = surv_wb_obj.active
surv_rows = surv_sheet_obj.max_row

cleaned_points =[]

#removes numbers from codes for checking
for i in range(1, surv_rows):
    cell_num_obj = surv_sheet_obj.cell(row = i, column = 1)
    cell_code_obj = surv_sheet_obj.cell(row = i, column = 5)
    cell_num = cell_num_obj.value
    cell_codes = cell_code_obj.value
    parsed_cell_codes = cell_codes.split()

    pattern = r'[0-9.]'
    for code in range(len(parsed_cell_codes)):
        un_numbered_code = re.sub(pattern, '', parsed_cell_codes[code])
        parsed_cell_codes[code] = un_numbered_code

    checked_parsed_cell_codes = []

    for code in range(len(parsed_cell_codes)):
        if (parsed_cell_codes[code] != ''):
            checked_parsed_cell_codes.append(parsed_cell_codes[code])
            
    checked_parsed_cell_codes.insert(0, cell_num)
    cleaned_points.append(checked_parsed_cell_codes)

points_with_weird_codes = []

#checks clean points against point list(s)
for i in cleaned_points:
    for code in range(1, len(i)):
        code_for_check = i[code]
        is_present = False
        for check_code in codes:
            if (code_for_check == check_code):
                is_present = True
        
        if (is_present == False):
            points_with_weird_codes.append(i[0])

#create and write to new sheet in survey file workbook,
#creates vlookups to show corresponding data in new sheet
surv_wb_obj.create_sheet('Review Points')
surv_sheet_review_points = surv_wb_obj["Review Points"]
for pt_num in range(len(points_with_weird_codes)):
    cell_A1 = surv_sheet_review_points.cell(row = pt_num + 1, column = 1)
    cell_A1.value = points_with_weird_codes[pt_num]

    cell_A2 = surv_sheet_review_points.cell(row = pt_num + 1, column = 2)
    cell_A2.value = f'=VLOOKUP(A{pt_num + 1},in!$A:$E, 2, FALSE)'

    cell_A3 = surv_sheet_review_points.cell(row = pt_num + 1, column = 3)
    cell_A3.value = f'=VLOOKUP(A{pt_num + 1},in!$A:$E, 3, FALSE)'

    cell_A4 = surv_sheet_review_points.cell(row = pt_num + 1, column = 4)
    cell_A4.value = f'=VLOOKUP(A{pt_num + 1},in!$A:$E, 4, FALSE)'

    cell_A5 = surv_sheet_review_points.cell(row = pt_num + 1, column = 5)
    cell_A5.value = f'=VLOOKUP(A{pt_num + 1},in!$A:$E, 5, FALSE)'
surv_wb_obj.save(path_surv)