import os
from zipfile import ZipFile
import shutil
import pandas
import logs
import numpy as np
import sys


def convert_xlsx(file_in, new_filename):
    if os.path.exists(file_in) is False:
        print(f"Файл {file_in} не найден")
        return False
    tmp_folder = 'tmp/convert_wrong_excel/'
    os.makedirs(tmp_folder, exist_ok=True)
    with ZipFile(file_in) as excel_container:
        excel_container.extractall(tmp_folder)
    wrong_file_path = os.path.join(tmp_folder, 'xl', 'SharedStrings.xml')
    correct_file_path = os.path.join(tmp_folder, 'xl', 'sharedStrings.xml')
    os.rename(wrong_file_path, correct_file_path)
    shutil.make_archive('tmp', 'zip', tmp_folder)
    os.replace('tmp.zip', new_filename)
    return True


def get_all_manuf():
    list_manuf = set()
    if os.path.exists('linkDict.txt') is False:
        logs.write_log('нет файла linkDict.txt с путем к справочнику')
        sys.exit()
    with open('linkDict.txt', 'r', encoding='utf-8') as path_dict:
        f_in = path_dict.readline().replace("\r", '').replace("\n", '')
    f_conv = "dict.xlsx"
    if convert_xlsx(f_in, f_conv) is False:
        logs.write_log(f'ошибка конвертации справочника коротких наименований')
        sys.exit()
    pd = pandas.read_excel(f_conv)
    pd.columns = ['manuf', 'good', 'name_search']
    select1 = pd[~pd.name_search.isna()].sort_values(by='manuf')
    list_manuf = select1.manuf.unique()
    return list_manuf

if __name__ == "__main__":
    get_all_manuf()
