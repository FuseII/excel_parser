"""Модуль предназначен для парсинга Excel файла. Информация считывается построчно.
Осуществляется переход по ссылкам внутри исходного документа, и если эти документы(таблицы) подходят (имеют нужную структуру)
 то также добавляем их данные в результат.
Результатом являются json файлы. Каждая строка исходного excel-документа - json файл.

Для вызова скрипта через командную строку (cmd):
parse_excel.exe 'путь_к_исходному_excel_файлу' 'опционально:директория_для_сохранения_результатов' 'опционально:директория_для_сохранения_excel_файлов'

Для преобразования в исполняемый файл:
pyinstaller --onefile parse_excel.py

Для вызова тестирования:
./parse_excel.py 'https://docs.google.com/spreadsheets/d/1uFpI6armr0XDzcxZKYC5mBAzgFcWqJcxt2cxkwt4aJ8/edit?gid=355897038#gid=355897038'
"""

# -*- coding: utf-8 -*-
import copy
import json
import sys

import pandas as pd
import requests
import re

from urllib.parse import urlparse, parse_qs
from pathlib import Path
from inner_parsers import parse_inner_excel_type_1, parse_inner_excel_type_2, parse_inner_excel_type_3, \
    data_preprocessing

INNER_PARSER_FUNCS = [parse_inner_excel_type_1, parse_inner_excel_type_2, parse_inner_excel_type_3]

# ссылка на файл для скачивания
LINK_SOURCE_FILE_PATH = r'https://docs.google.com/spreadsheets/d/1uFpI6armr0XDzcxZKYC5mBAzgFcWqJcxt2cxkwt4aJ8/edit?gid=355897038#gid=355897038'
# путь к исходному excel файлу по умолчанию
SOURCE_FILE_PATH = './data/Сводный перечень текущих работ для графиков.xlsx'
# путь к директории для сохранения всех внутренних excel-файлов по умолчанию
SAVED_ALL_EXCEL_DIRECTORY_PATH = './downloaded_all'
# путь к директории для сохранения результатов по умолчанию
SAVED_RESULTS_PATH = './saved_jsons'

COLUMNS_TO_BE_MATCHED = ['Разработка_материалов_существующего_положения_эскизы',
                         'Рассмотрение_эскизов_выбор_рекомендуемого_варианта',
                         'Разработка_обосновывающих_материалов_проектного_предложения',
                         'Расмотрение_обосновывающих_материалов_получение_замечаний',
                         'Корректировка_обосновывающих_материалов_по_замечаниям_ОИВ_повторная_рассылка',
                         'Корректировка_материалов_по_замечаниям_ОИВ_Подготовка_и_согласование_макета_и_НПМ_получение_согласований_ОИВ',
                         'Подготовка_презентации_на_ГЗК',
                         'Проведение_РГ_ГЗК_И_ГЗК',
                         'Подготовка_Утв_Части',
                         'ПУ_МКА',
                         'АИС_СД_УКД',
                         'АИС_СД_ОИВ',
                         'Доработка_материалов_по_замч_ОИВ',
                         'ПУ_ПМ',
                         'ОАУ_ПМ']


def download_file_stream(url, local_filename):
    """
    Скачивает файл по ссылке
    :param url: ссылка
    :param local_filename: имя сохраняемого файла
    :return:
    """
    try:
        with requests.get(url, stream=True) as r:
            r.raise_for_status()
            with open(local_filename, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
    except Exception as e:
        print(f"Ошибка загрузки фала {e}")
        return False
    print(f'Файл успешно сохранён как {local_filename}')
    return True


def check_excel(local_filename):
    """
    Прооверяет, что это подходящий excel документ
    :param: local_filename - путь к файлу
    """
    # проверяем что такой путь существует
    if not Path(local_filename).exists():
        print(f"Путь {local_filename} не существует!")
        return False
    df = pd.ExcelFile(local_filename)
    df = df.parse(df.sheet_names[0])
    df = df.astype(str)
    df.dropna(axis=0, how='all', inplace=True)
    df = df.iloc[3:]
    df.columns = df.iloc[0]
    # df.columns = df.columns.str.replace('\n', '', regex=False)
    # df.columns = df.columns.str.replace(' ', '_', regex=False)
    # df.columns = df.columns.str.strip()
    # df = df.replace("\n", ' ', regex=False)
    # df = df.replace("\t", ' ', regex=False)
    df = df.drop(3).reset_index(drop=True)
    df = data_preprocessing(df)
    # кол-во совпавших имён колонок
    matched_columns = 0
    for column in COLUMNS_TO_BE_MATCHED:
        if column in df.columns:
            matched_columns += 1
    # если все колонки есть в таблице
    if matched_columns >= len(COLUMNS_TO_BE_MATCHED):
        return True
    else:
        return False


def convert_google_doc_to_download_link(url):
    """
        Преобразует ссылку Google Sheets в ссылку для скачивания файла в формате XLSX.
        Параметры:
        url (str): ссылка на Google Sheets документ.
        Возвращает:
        str или None: ссылка для скачивания формата XLSX или None, если ссылка невалидная.
    """
    # Попытка извлечь ID документа
    match = re.search(r"https://docs\.google\.com/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if not match:
        return None
    doc_id = match.group(1)
    # Парсим gid из параметров (если есть)
    parsed_url = urlparse(url)
    query_params = parse_qs(parsed_url.query)
    gid = None
    if 'gid' in query_params:
        gid = query_params['gid'][0]
    else:
        # Иногда gid может быть в хеш-части URL
        if parsed_url.fragment.startswith("gid="):
            gid = parsed_url.fragment.split('=')[1]
    if gid:
        download_link = f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=xlsx&gid={gid}"
    else:
        # Если gid не найден, создаём ссылку без него (скачается первая таблица)
        download_link = f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=xlsx"
    return download_link


def parse_source_excel(list_of_dicts):
    """
   Основная функция для парсинга.
   :param list_of_dicts: исходный список словарей из исходного excel-файла
    return: None
    """
    # словарь соответствия 'Наименование_работы' -> 'Шифр_работы'. Нужно для постобработки
    work_name_to_code_dict = work_name_to_code(list_of_dicts)
    # список посещённых сслыок
    visited_links = []
    # кол-во json-документов
    json_number = 0
    for obj in list_of_dicts:
        json_number += 1
        obj_result = copy.deepcopy(obj)
        obj_result["data_from_links"] = dict()
        for k, v in obj.items():
            if isinstance(v, str):  # определяем ссылки
                # выделяем и проходимся по ссылкам
                url_pattern = r"https?://[^\s]+"
                links = re.findall(url_pattern, v)

                # проход по ссылкам и закачка excel-файлов
                for link in links:
                    link_for_download = convert_google_doc_to_download_link(link)
                    if link_for_download is not None:
                        # print(link_for_download)
                        if link_for_download not in visited_links:
                            visited_links.append(link_for_download)
                            link_number = visited_links.index(link_for_download)
                            local_filename = f'{SAVED_ALL_EXCEL_DIRECTORY_PATH}/{link_number}.xlsx'
                            # local_filename = f'./downloaded_all/{link_number}.xlsx'
                            downloaded = download_file_stream(url=link_for_download, local_filename=local_filename)
                            if not downloaded:
                                continue
                        else:
                            link_number = visited_links.index(link_for_download)
                            local_filename = f'{SAVED_ALL_EXCEL_DIRECTORY_PATH}/{link_number}.xlsx'
                            # local_filename = f'./downloaded_all/{link_number}.xlsx'

                        # пытаемя подобрать правильную функцию для парсинга файла
                        for parser_func_type, func in enumerate(INNER_PARSER_FUNCS):
                            try:
                                parsed_data = func(local_filename)
                                if parsed_data is None:
                                    continue
                                obj_result_copy = copy.deepcopy(obj_result)
                                obj_result_copy["data_from_links"].update({k: parsed_data})
                                obj_result_copy = postprocessing_json(obj_result_copy, work_name_to_code_dict)
                                obj_result_copy['Вариант_парсера'] = str(parser_func_type + 1)
                                obj_result = obj_result_copy
                                with open(f'{SAVED_RESULTS_PATH}/{json_number}.json', 'w', encoding='utf-8') as f:
                                    json.dump(obj_result, f, ensure_ascii=False, indent=4)
                                print(f"Результат парсинга сохранён в {SAVED_RESULTS_PATH}/{json_number}")
                                break
                            except Exception as ex:
                                # print(f"file_number = {link_number}, exception = {ex}, parser = {parser_func_type}")
                                pass
                        # если внутренний excel-файл подходит, то выполняем предобработку и сохраняем его в директорию
                        # if check_excel(local_filename):
                        #     parsed_data = parse_inner_excel_type_1(local_filename)
                        #     obj_result["data_from_links"].update({k: parsed_data})
                        #     obj_result = postprocessing_json(obj_result, work_name_to_code_dict)
                        #     with open(f'{SAVED_RESULTS_PATH}/{json_number}.json', 'w', encoding='utf-8') as f:
                        #         json.dump(obj_result, f, ensure_ascii=False, indent=4)


def postprocessing_json(obj, work_name_to_code_dict):
    """
    Дополнительная обработка получившегося json файла
    :param obj: входящий json
    :param work_name_to_code_dict:  Сопоставление 'Наименование_работы' -> 'Шифр_работы'.
    :return:
    """
    # добавим шифр работы в каждую работу в data_from_links из глобального excel-файла
    # print("in postprocessing")
    # print(json.dumps(obj, ensure_ascii=False, indent=4))
    # print(json.dumps(work_name_to_code_dict, ensure_ascii=False, indent=4))
    data_from_links = obj["data_from_links"]
    for link in data_from_links:
        new_work_list = list()
        for work in data_from_links[link]:
            new_work = dict()
            new_work["Номер"] = work.get("Номер", "None")
            new_work["Шифр_работы"] = work_name_to_code_dict.get(work.get("Наименование_работы", "None"), "None")
            # new_work["Шифр_работы"] = obj["Шифр_работы"]
            for key, item in work.items():
                new_work[key] = work[key]
            new_work_list.append(new_work)
        data_from_links[link] = new_work_list

    return obj


def work_name_to_code(list_of_dicts):
    """
    Сопоставление 'Наименование_работы' -> 'Шифр_работы'.
    :return:dict
    """
    result = dict()
    for obj in list_of_dicts:
        result[obj['Наименование_проекта']] = obj.get('Шифр_работы', "None")
    return result


# https://docs.google.com/spreadsheets/d/1uFpI6armr0XDzcxZKYC5mBAzgFcWqJcxt2cxkwt4aJ8/edit?gid=355897038#gid=355897038

def get_source_excel(link, download_path):
    """Скачивает исходный excel-файл в директорию."""

    link = convert_google_doc_to_download_link(link)
    download_file_stream(url=link, local_filename=download_path)


def is_link(text):
    """Проверяет ссылка ли это"""
    url_pattern = r"https?://[^\s]+"
    links = re.findall(url_pattern, text)
    if len(links) != 0:
        return True
    else:
        return False


def main():
    global SOURCE_FILE_PATH
    global SAVED_RESULTS_PATH
    global SAVED_ALL_EXCEL_DIRECTORY_PATH

    current_dir = Path.cwd()
    if is_link(sys.argv[1]):
        link = sys.argv[1]
        SOURCE_FILE_PATH = current_dir.joinpath('source')
        SOURCE_FILE_PATH.mkdir(parents=True, exist_ok=True)
        SOURCE_FILE_PATH = f"{SOURCE_FILE_PATH}/source_excel.xlsx"
        get_source_excel(link, SOURCE_FILE_PATH)
    else:
        SOURCE_FILE_PATH = sys.argv[1]

    SAVED_RESULTS_PATH = current_dir.joinpath('saved_jsons')
    SAVED_RESULTS_PATH.mkdir(parents=True, exist_ok=True)

    SAVED_ALL_EXCEL_DIRECTORY_PATH = current_dir.joinpath('downloaded_all')
    SAVED_ALL_EXCEL_DIRECTORY_PATH.mkdir(parents=True, exist_ok=True)

    df = pd.ExcelFile(SOURCE_FILE_PATH)
    # имена страниц внутри excel-файла
    sheet_names = df.sheet_names
    # рассматриваем таблицу только с 0-го листа
    df = df.parse(sheet_names[0])
    # предобработка датафрейма
    df = data_preprocessing(df)
    # преобразуем каждую строчку датафрейма в словарь и добавляем её в список
    list_of_dicts = df.to_dict(orient='records')
    # вызываем основную функцию
    parse_source_excel(list_of_dicts)
    print("Программа завершена.")


def test():
    # local_file_path = './data/ППТ ТПУ Петров С.А..xlsx'
    # print(parse_inner_excel_type_1(local_filename=local_file_path))
    # local_file_path = './data/108.xlsx'
    # parsed_inner_result = parse_inner_excel_type_2(local_filename=local_file_path)
    local_file_path = './data/100.xlsx'
    parsed_inner_result = parse_inner_excel_type_2(local_filename=local_file_path)
    print(json.dumps(parsed_inner_result, ensure_ascii=False, indent=4))
    # obj_result = dict()
    # obj_result["data_from_links"] = dict()
    # obj_result["data_from_links"].update({'k': parsed_inner_result})
    # obj_result = postprocessing_json(obj_result, work_name_to_code_dict={})
    # print(json.dumps(obj_result, ensure_ascii=False, indent=4))


if __name__ == '__main__':
    main()
    # test()
    # df = pd.ExcelFile(SOURCE_FILE_PATH)
    # имена страниц внутри excel-файла
    # sheet_names = df.sheet_names
    # рассматриваем таблицу только с 0-го листа
    # df = df.parse(sheet_names[0])
    # предобработка датафрейма
    # df = data_preprocessing(df)
    # преобразуем каждую строчку датафрейма в словарь и добавляем её в список
    # list_of_dicts = df.to_dict(orient='records')
    # вызываем основную функцию
    # parse_source_excel(list_of_dicts)
    # print("Программа завершена.")
    # test()
    # main()
