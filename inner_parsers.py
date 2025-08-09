import pandas as pd
from pandas import DataFrame
import json

COLUMNS_TO_BE_MATCHED_TYPE_1 = ['Разработка_материалов_существующего_положения_эскизы',
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

COLUMNS_TO_BE_MATCHED_TYPE_2 = ['Направление_заявки_на_презентацию',
                                'Подготовка_презентации_выявление_ограничений',
                                'Рассмотрение_и_согласование_презентационных_материалов_на_предмет_рисков_подготовка_ТЗ_и_границ_для_заказа_ИД',
                                'Подготовка_и_прохождение_ГЗК',
                                'подготовка_направление_распоряжения_и_заявки_запрос_исходных_данных_и_геоподосновы_М_1500',
                                '1_Разработка_материалов_существующего_положения_эскизы_вариантов_планировочного_решения',
                                'Рассмотрение_эскизов_вариантов_планировочного_решения_выбор_рекомендуемого_варианта_(с_учётом_предложений_АО_"Мосинжпроект")',
                                '2_Разработка_обосновывающих_материалов_проектного_предложения',
                                'Расмотрение_обосновывающих_материалов_выдача_замечаний',
                                'Корректировка_материалов_по_замечаниям_ОИВ_получение_согласований_ОИВ',
                                'Дополнительное_согласование_с_РЖД', 'подготовка_презентации_на_ГЗК',
                                'Проведение_РГ_ГЗК_И_ГЗК', 'подготовка_Утв_Части',
                                'Проверка_УЧ_в_кур_Управлении_и_ПУ_МКА',
                                'Устранение_замеч_ПУ_МКА_загрузка_в_АИС_СД', 'АИС_СД_УКД', 'АИС_СД_ОИВ',
                                'Согласование_с_Росимуществом_и_ФАЖТ', 'Согласование_с_Мособластью',
                                'Прохождение_МГЭ', 'Доработка_материалов_по_замч_ОИВ',
                                'проверка_в_ПУ_МКА',
                                'Прохождение_ППМ_корректировка_по_замеч_повторное_рассмотрение',
                                'ОАУ_ПМ']


def data_preprocessing(df: DataFrame) -> DataFrame:
    """
    Предобработка датафрейма. Удаление пустых строк, ненужных символов и т.д.
    :param df:DataFrame
    :return:DataFrame
    """
    # удалим строки, в которых все значения NaN
    df.dropna(axis=0, how='all', inplace=True)
    # преобразуем все колонки к строковому типу
    df = df.astype(str)

    df = df.replace({'nan': 'None', 'NaT': 'None'})
    df = df.map(lambda x: x.strip())

    df.columns = df.columns.str.strip()
    df.columns = df.columns.str.replace(r'\.', '', regex=True)
    df.columns = df.columns.str.replace(r'\,', '', regex=True)
    df.columns = df.columns.str.replace(r'\:', '', regex=True)
    df.columns = df.columns.str.replace('/', '_', regex=True)
    df = df.replace('"', '', regex=True)
    df = df.replace('\\', '', regex=False)
    df = df.replace('\\"', '', regex=False)

    df = df.replace('+', '', regex=False)
    df.columns = df.columns.str.replace('+', '', regex=False)

    df = df.replace(r'[\r\n\t]+', ' ', regex=True)
    df.columns = df.columns.str.replace(r'[\r\n\t]+', ' ', regex=True)

    # замена нескольких идущих подряд пробелов - одним
    df = df.replace('\\s+', ' ', regex=True)
    df.columns = df.columns.str.replace('\\s+', '_', regex=True)

    # замена нескольких идущих подряд символов '_' - одним
    df.columns = df.columns.str.replace(r'_{2,}', '_', regex=True)

    df.columns = df.columns.str.replace('№', 'Номер', regex=False)

    return df


def parse_inner_excel_type_1(local_filename: str):
    """
    Парсинг полученного по ссылке excel-файла (внутреннего)
    :param: local_filename:str - путь к файлу
    return: [dict()]
    """
    inner_df = pd.ExcelFile(local_filename)
    inner_df = inner_df.parse(inner_df.sheet_names[0])
    inner_df = inner_df.astype(str)
    inner_df.dropna(axis=0, how='all', inplace=True)
    inner_df = inner_df.iloc[3:]
    inner_df.columns = inner_df.iloc[0]
    inner_df = inner_df.drop(3).reset_index(drop=True)
    inner_df.columns.values[3] = 'Дата'
    inner_df.columns.values[-1] = 'Статус'
    inner_df = data_preprocessing(inner_df)
    result_list = []
    i = 0
    while i < len(inner_df):
        # стоит ли брать одну ли две строчки
        if (inner_df["Наименование_работы"][i] != 'None') and (inner_df["Наименование_работы"][i + 1] == 'None'):
            step = 2
        else:
            step = 1
        part_df = inner_df.iloc[i:i + step]  # Получаем одну/две строки
        result = part_df.iloc[0, :3].to_dict()
        result.update(part_df[["Статус"]].iloc[[0]].to_dict(orient='records')[0])
        i += step
        part_df.loc[:, 'Дата'] = part_df['Дата'].str.replace('\\s+', '_', regex=True)
        part_df = part_df.set_index('Дата').iloc[:, 3:-1]

        part_df = part_df.to_dict()
        result.update(part_df)
        result_list.append(result)
    matched = 0
    for column in inner_df.columns:
        if (column != 'nan') and (column in COLUMNS_TO_BE_MATCHED_TYPE_1):
            matched = matched + 1
    if matched < len(inner_df.columns)//2:
        return None

    return result_list


def parse_inner_excel_type_2(local_filename: str, start_string:int = 8):
    """
    Парсинг полученного по ссылке excel-файла (внутреннего)
    :param: local_filename:str - путь к файлу
    return: [dict()]
    """
    inner_df = pd.ExcelFile(local_filename)
    inner_df = inner_df.parse(inner_df.sheet_names[0])
    inner_df = inner_df.astype(str)
    inner_df.dropna(axis=0, how='all', inplace=True)
    inner_df = inner_df.iloc[start_string:start_string+4]
    inner_df.columns = inner_df.iloc[0]
    inner_df = inner_df.drop(start_string).reset_index(drop=True)
    inner_df = data_preprocessing(inner_df)
    inner_df.set_index('nan')
    data = inner_df.to_dict('list')
    # result = inner_df.to_dict('records')
    data['Даты'] = data.pop('nan')
    result = dict()
    dates = data['Даты']
    data.pop('Даты')
    # print(dates)
    result_list = []
    for key, value in data.items():
        key = key.replace(r'"', '')
        result[key] = dict()
        for i, date in enumerate(dates):
            date_replaced = date.replace(' ', '_')  # !
            result[key][date_replaced] = value[i]
    # print(json.dumps(obj=result, indent=4, ensure_ascii=False))
    result_list.append(result)
    matched = 0
    for column in inner_df.columns:
        if (column != 'nan') and (column in COLUMNS_TO_BE_MATCHED_TYPE_2):
            matched = matched + 1
    if matched < len(inner_df.columns)//2:
        return None
    return result_list

def parse_inner_excel_type_3(local_filename: str):
    return parse_inner_excel_type_2(local_filename, start_string=13)


def main():
    result  = parse_inner_excel_type_2(local_filename='./data/106.xlsx')
    print(result)


if __name__ == '__main__':
    main()