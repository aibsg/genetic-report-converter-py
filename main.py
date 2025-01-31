from converter import *

import os
from converter import process_excel_iterative

def find_and_process_files(base_dir, output_file):
    main_counter = 1
    for root, dirs, files in os.walk(base_dir):
        if 'ДНК' in dirs:
            dnk_dir = os.path.join(root, 'ДНК')
            for dnk_root, dnk_dirs, dnk_files in os.walk(dnk_dir):
                for file in dnk_files:
                    if file.startswith('Заключение') and file.endswith('.xlsx'):
                        input_file = os.path.join(dnk_root, file)
                        print(f"Processing file: {input_file}")
                        main_counter = process_excel_iterative(input_file, output_file, main_counter)
                        break
                break
        else:
            for file in files:
                if file.startswith('Заключение') and file.endswith('.xlsx'):
                    input_file = os.path.join(root, file)
                    print(f"Processing file: {input_file}")
                    main_counter = process_excel_iterative(input_file, output_file, main_counter)

if __name__ == "__main__":
    base_dir = r"C:\Users\gglol\Documents\ДНК Договора 2024 тестовый"
    output_file = r"C:\Users\gglol\Documents\ДНК Договора 2024 тестовый\Приложение 2.xlsx"
    find_and_process_files(base_dir, output_file)