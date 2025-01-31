import openpyxl
from datetime import datetime
import re

class TestPoint:
    def __init__(self, test_date, org_name, sex, ind_num, name, birth_date, locus_dic, conclusion):
        self.test_date = test_date
        self.org_name = org_name
        self.sex = sex
        self.ind_num = ind_num
        self.name = name
        self.birth_date = birth_date
        self.locus_dic = locus_dic
        self.conclusion = conclusion

def create_dic(sheet, row, locus_count, start_locus_col, locus_head_row):
    res_dic = {}
    for col in range(start_locus_col, start_locus_col + locus_count):
        res_dic[sheet.cell(locus_head_row, col).value] = [sheet.cell(row, col).value, sheet.cell(row + 1, col).value]
    return res_dic

def conclusion_calculate(input_string):
    if input_string is None:
        return "Ошибка: Невозможно обработать входную строку"
    
    input_string = input_string.lower()  # ignore case

    if "отец соответствует" in input_string and "мать не тестирована" in input_string:
        return "да/нет"
    elif "родители не соответствуют" in input_string:
        return "нет/нет"
    elif "отец соответствует" in input_string and "мать соответствует" in input_string:
        return "да/да"
    elif "отец не соответствует" in input_string or "отец не тестирован" in input_string and "мать не тестирована" in input_string or "мать не соответствует" in input_string:
        return "нет/нет"
    elif "отец не соответствует" in input_string or "отец не тестирован" in input_string and "мать соответствует" in input_string:
        return "нет/да"
    elif "родители соответствуют" in input_string or 'родители сответствуют' in input_string:
        return "да/да"
    elif "получен микросателлитный профиль" in input_string:
        return "тест"
    elif "получен микросател- литный профиль" in input_string:
        return "тест"
    else:
        return input_string

def write(test_point, output_sheet, counter, main_counter, output_locus_names, output_header_row, output_locus_start_col, output_locus_end_col):
    output_sheet.cell(main_counter + counter + output_header_row, 1, main_counter + counter)
    output_sheet.cell(main_counter + counter + output_header_row, 2, test_point.org_name)
    output_sheet.cell(main_counter + counter + output_header_row, 3, test_point.test_date)
    output_sheet.cell(main_counter + counter + output_header_row, 5, "F")
    if test_point.sex == "п":
        output_sheet.cell(main_counter + counter + output_header_row, 6, "F1")
    else:
        output_sheet.cell(main_counter + counter + output_header_row, 6, "M")
    output_sheet.cell(main_counter + counter + output_header_row, 7, test_point.ind_num)
    output_sheet.cell(main_counter + counter + output_header_row, 8, test_point.name)
    output_sheet.cell(main_counter + counter + output_header_row, 9, test_point.birth_date)
    for i, locus in enumerate(output_locus_names):
        if locus in test_point.locus_dic:
            output_sheet.cell(main_counter + counter + output_header_row, output_locus_start_col + i * 2, test_point.locus_dic[locus][0])
            output_sheet.cell(main_counter + counter + output_header_row, output_locus_start_col + i * 2 + 1, test_point.locus_dic[locus][1])
        else:
            output_sheet.cell(main_counter + counter + output_header_row, output_locus_start_col + i * 2, "")
            output_sheet.cell(main_counter + counter + output_header_row, output_locus_start_col + i * 2 + 1, "")
    output_sheet.cell(main_counter + counter + output_header_row, output_locus_end_col + 1, test_point.conclusion)

def process_excel_iterative(input_file, output_file, main_counter, config):
    counter = 0
    input_wb = openpyxl.load_workbook(input_file, data_only=True)
    input_sheet = input_wb.active

    output_wb = openpyxl.load_workbook(output_file)
    output_sheet = output_wb.active

    org_name = input_sheet.cell(config['name'][0], config['name'][1]).value
    testing_date = input_sheet.cell(config['date'][0], config['date'][1]).value

    output_locus_names = [output_sheet.cell(config['output_header_row'], col).value for col in range(config['output_locus_start_col'], config['output_locus_end_col'], 2)]
    col = config['start_locus_col']
    locus_count = 0
    while input_sheet.cell(config['locus_head_row'], col).value != None:
        locus_count += 1    
        col += 1
    for row in range(config['start_data_row'], input_sheet.max_row + 1, 2):
        sex = input_sheet.cell(row, 2).value
        cell_value = input_sheet.cell(row, 1).value
        if sex is None:
            sex = "м"
        if cell_value is None or not re.match(r'^\d+/24 *днк$', cell_value.lower()) or sex not in "пм":
            continue
        locus_dic = create_dic(input_sheet, row, locus_count, config['start_locus_col'], config['locus_head_row'])
        input_end_locus_col = len(locus_dic.keys()) + config['start_locus_col'] - 1
        test_name = input_sheet.cell(row, 3).value
        ind_num = input_sheet.cell(row, 4).value
        birth_date = input_sheet.cell(row, 5).value
        work_cell_con = 1
        if input_sheet.cell(row, input_end_locus_col + work_cell_con).value == None:
            work_cell_con += 1
        con_val = input_sheet.cell(row, input_end_locus_col + work_cell_con).value
        if con_val == None:
            conclusion = ""  # Надо исправить потом
        else:
            conclusion = conclusion_calculate(con_val)
        test_point = TestPoint(testing_date, org_name, sex, ind_num, test_name, birth_date, locus_dic, conclusion)
        write(test_point, output_sheet, counter, main_counter, output_locus_names, config['output_header_row'], config['output_locus_start_col'], config['output_locus_end_col'])
        counter += 1

    output_wb.save(output_file)
    return main_counter + counter