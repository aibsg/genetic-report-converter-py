import os
from converter import process_excel_iterative

def find_and_process_files(base_dir, output_file, main_counter, config):
    for root, dirs, files in os.walk(base_dir):
        if 'ДНК' in dirs:
            dnk_dir = os.path.join(root, 'ДНК')
            for dnk_root, dnk_dirs, dnk_files in os.walk(dnk_dir):
                for file in dnk_files:
                    if file.startswith('Заключение') and file.endswith('.xlsx'):
                        input_file = os.path.join(dnk_root, file)
                        print(f"Processing file: {input_file}")
                        main_counter = process_excel_iterative(input_file, output_file, main_counter, config)
    return main_counter

if __name__ == "__main__":
    base_dir = r"C:\Users\gglol\Documents\ДНК Договора 2024 тестовый"
    output_file = r"C:\Users\gglol\Documents\ДНК Договора 2024 тестовый\Приложение 2.xlsx"
    config = {
        'name': (3, 5),
        'date': (4, 5),
        'start_header_row': 5,
        'locus_head_row': 6,
        'start_locus_col': 6,
        'start_data_row': 7,
        'output_header_row': 5,
        'output_locus_start_col': 10,
        'output_locus_end_col': 41
    }
    main_counter = 0
    main_counter = find_and_process_files(base_dir, output_file, main_counter, config)