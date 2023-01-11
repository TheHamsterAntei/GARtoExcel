import numpy
import numpy as np
import zipfile
import os
import pandas as pd
import xml.etree.ElementTree as ET
import sys
from tkinter import messagebox
from tkinter import ttk
import tkinter as tk
import logging
import time
#необходим XlsxWriter


#Константы
#Типы данных в таблицах pandas
DICT_DF_DTYPES = {
    'AO_GUID': 'string',
    'AO_LEVEL': 'int16',
    'AO_HOUSENUM': 'string',
    'AO_KADASTNUM': 'string',
    'AO_POSTINDEX': 'string',
    'AO_OKATO': 'string',
    'AO_OKTMO': 'string',
    'AO_IFNSFL': 'string',
    'AO_IFNSUL': 'string',
    'AO_TERRIFNSFL': 'string',
    'AO_TERRIFNSUL': 'string',
    'AO_REESTRNUM': 'string'
}
#Краткие наименования типов домов
DICT_HOUSETYPES = {
    '1': 'влд.',
    '2': 'д.',
    '3': 'двлд.',
    '4': 'г-ж',
    '5': 'зд.',
    '6': 'шахта',
    '7': 'стр.',
    '8': 'соор.',
    '9': 'литера',
    '10': 'к.',
    '11': 'подв.',
    '12': 'кот.',
    '13': 'п-б',
    '14': 'ОНС'
}
#Соответствие параметров базы GAR столбцам pandas
DICT_PARAMS = {
    '1': 10,
    '2': 11,
    '3': 12,
    '4': 13,
    '5': 7,
    '6': 8,
    '7': 9,
    '8': 6,
    '13': 17,
    '16': 15
}
#Названия excel-таблиц по уровням объектов
DICT_LEVELS = {
    1: 'subjects',
    2: 'adm_districts',
    3: 'mun_districts',
    4: 'settlements',
    5: 'towns',
    6: 'communities',
    7: 'territories',
    8: 'streets',
    9: 'sub_territories',
    10: 'houses'
}

#Глобальные переменные
root = tk.Tk()
root.geometry('400x280')
root.title('Обработка БД ГАР')
progress_main = 0
progress_add = 0
label_value = 'Подготовка...'
subject_num = 0
label_subj_value = str('В обработке субъект №' + str(subject_num))
destroyed = False
memory_usage = 0
global_start_time = time.time()
progress_log = 0
progress_log_time = global_start_time
read_levels = [1, 5, 6, 8]
houses_per_excel = 50000

#Отрисовка объектов в графическом окне
pb_main = ttk.Progressbar(
    root,
    orient='horizontal',
    mode='determinate',
    length=360
)
pb_main.grid(column=0, row=0, columnspan=2, padx=20, pady=20)
label = ttk.Label(root, text=label_value)
label.grid(column=0, row=1, columnspan=2)
pb_add = ttk.Progressbar(
    root,
    orient='horizontal',
    mode='determinate',
    length=360
)
pb_add.grid(column=0, row=2, columnspan=2, pady=20)
label_subj = ttk.Label(root, text=label_subj_value)
label_subj.grid(column=0, row=3, columnspan=2, pady=20)
label_usage = ttk.Label(root, text='Использовано 0 б', anchor=tk.CENTER)
label_usage.grid(column=0, row=4)
label_time = ttk.Label(root, text='Прошло 0 сек.')
label_time.grid(column=1, row=4, columnspan=1)
label_prog = ttk.Label(root, text='Нет параметра обработки')
label_prog.grid(column=0, row=5, columnspan=2, pady=20)


def close_window():
    global destroyed
    if messagebox.askokcancel("Выход", "Завершить выполнение программы?"):
        destroyed = True
        root.destroy()

root.protocol("WM_DELETE_WINDOW", close_window)


#Найти в архиве файл по субъекту и фрагменту названия
def findInZip(archive, subject, filename):
    namelist = archive.namelist()
    region = str(subject)
    if len(region) == 1:
        region = '0' + region
    for i in namelist:
        if i[0:2] == region:
            if i[3:3+len(filename)] == filename:
                return i
            else:
                continue
        else:
            continue
    return None


#Обновить интерфейс
def update_bars(first, second, log, object, label_new):
    global destroyed
    if destroyed:
        sys.exit(0)
    global label_value, progress_main, progress_add, pb_main, pb_add, root, label, label_usage, label_time
    global label_prog, progress_log, progress_log_time
    label_value = label_new
    progress_main = first
    progress_add = second
    pb_main['value'] = progress_main
    pb_add['value'] = progress_add
    label['text'] = label_value
    label_usage['text'] = 'Занято: ' + return_memory_usage_as_str()
    new_time = time.time()
    label_time['text'] = 'Прошло ' + return_time_as_str(new_time - global_start_time)
    label_prog['text'] = ('Скорость обработки: ' + return_factor_speed_as_str((log - progress_log) /
                                                                              (new_time - progress_log_time)) +
                          ' ' + object + 'ов в сек.')
    progress_log = log
    progress_log_time = new_time - 0.00001
    root.update()


#Вернуть скорость обработки строкой
def return_factor_speed_as_str(log):
    if log >= 100.0:
        return str(int(log))
    else:
        if 1.0 <= log < 100.0:
            return str(int(log * 10) / 10)
        else:
            return str(int(log * 100) / 100)


#Вернуть использование памяти строкой
def return_memory_usage_as_str():
    pretype = 'б'
    if memory_usage < 1024:
        usage = memory_usage
        return str(usage) + '.0 ' + pretype
    if 1024 <= memory_usage < 1024*1024:
        usage = int(memory_usage / 102.4)
        pretype = 'кб'
        return str(int(usage / 10)) + '.' + str(usage % 10) + ' ' + pretype
    if 1024*1024 <= memory_usage:
        usage = int(memory_usage / (1024*102.4))
        pretype = 'мб'
        return str(int(usage / 10)) + '.' + str(usage % 10) + ' ' + pretype


#Вернуть время строкой
def return_time_as_str(timer):
    if timer > 1:
        timer = int(timer)
        if timer < 60:
            return str(timer) + ' сек.'
        if 60 <= timer < 60 * 60:
            if timer % 60 == 0:
                return str(int(timer / 60)) + ' мин.'
            else:
                return str(int(timer / 60)) + ' мин. ' + str(timer % 60) + ' сек.'
        if 60 * 60 <= timer < 60 * 60 * 24:
            timer = int(timer / 60)
            if timer % 60 == 0:
                return str(int(timer / 60)) + ' ч.'
            else:
                return str(int(timer / 60)) + ' ч. ' + str(timer % 60) + ' мин.'
        if 60 * 60 * 24 <= timer:
            timer = int(timer / (60 * 60))
            if timer % 24 == 0:
                return str(int(timer / 24)) + ' дн.'
            else:
                return str(int(timer / 24)) + ' дн. ' + str(timer % 24) + ' ч.'
    else:
        return str(int(timer * 1000)) + ' мс.'


#Вычислить потребление памяти объектом
def deep_getsizeof(obj):
    size_of = sys.getsizeof
    if isinstance(obj, dict):
        return size_of(obj) + sum(size_of(i) + size_of(j) for i, j in obj.items()) + sum(size_of(i) for i in obj.values())
    if isinstance(obj, list):
        if len(obj) == 0:
            return size_of(obj)
        else:
            if not isinstance(obj[0], pd.DataFrame):
                return size_of(obj) + sum([deep_getsizeof(i) for i in obj])
            else:
                return size_of(obj) + sum([i.memory_usage(index=False, deep=True).sum() for i in obj])
    if isinstance(obj, pd.DataFrame):
        return obj.memory_usage(index=False, deep=True).sum()
    return size_of(obj)


#Бинарный поиск в списке строк вида ("id:№ строки в pandas")
def binary_found_in_obj_list(lst, obj_id):
    start = 0
    end = len(lst) - 1

    if start > end:
        return None

    while end >= start:
        middle = (start + end) // 2
        temp_obj = lst[middle].split('=')
        value = int(temp_obj[0])
        if obj_id < value:
            end = middle - 1
            continue
        if obj_id > value:
            start = middle + 1
            continue
        if obj_id == value:
            return int(temp_obj[1])
    return None


#Чтение региона
def read_subject(n):
    global memory_usage, progress_log
    step_size = houses_per_excel
    total_log = 0
    subj_time = time.time()
    update_bars(0, 0, 0, 'файл', 'Поиск файлов в архиве...')
    archive = zipfile.ZipFile('GAR\\gar_xml.zip')
    found = findInZip(archive, n, 'AS_ADDR_OBJ')
    params = findInZip(archive, n, 'AS_ADDR_OBJ_PARAMS')
    hierarchy = findInZip(archive, n, 'AS_ADM_HIERARCHY')
    mun_hierarchy = findInZip(archive, n, 'AS_MUN_HIERARCHY')
    houses = findInZip(archive, n, 'AS_HOUSES')
    houses_params = findInZip(archive, n, 'AS_HOUSES_PARAMS')
    if found and params and hierarchy and mun_hierarchy and houses and houses_params:
        if not os.path.exists('Data\\' + str(n)):
            os.mkdir('Data\\' + str(n))
        update_bars(0, 50, 5, 'файл', 'Подготовка файлов...')
        progress_log = 0
        archived_main = archive.open(found)
        archived_params = archive.open(params)
        archived_hierarchy = archive.open(hierarchy)
        archived_mun_hierarchy = archive.open(mun_hierarchy)
        archived_houses = archive.open(houses)
        archived_houses_params = archive.open(houses_params)
        update_bars(1, 0, 5, 'файл', 'Обработка адресных объектов...')
        progress_log = 0
        df_main = pd.DataFrame({
            'NAME': [], #0
            'AO_ID': [], #1
            'AO_GUID': [], #2
            'AO_LEVEL': [], #3
            'AO_PARENT': [], #4
            'AO_HOUSENUM': [], #5
            'AO_KADASTNUM': [], #6
            'AO_POSTINDEX': [], #7
            'AO_OKATO': [], #8
            'AO_OKTMO': [], #9
            'AO_IFNSFL': [], #10
            'AO_IFNSUL': [], #11
            'AO_TERRIFNSFL': [], #12
            'AO_TERRIFNSUL': [], #13
            'AO_FORMALNAME': [], #14
            'AO_OFFICIALNAME': [], #15
            'AO_SHORTNAME': [], #16
            'AO_REESTRNUM': [], #17
            'AO_ADMADDRESS': [], #18
            'AO_MUNADDRESS': [], #19
            'DETAIL': [], #20
            'FILE': [], #21
        }).astype({'FILE': 'int16'})
        df_aos = pd.DataFrame({
            'AO_ID': [],
            'AO_GUID': [],
            'AO_LEVEL': [],
            'AO_FORMALNAME': [],
            'AO_SHORTNAME': []
        })
        df_houses = pd.DataFrame({
            'AO_ID': [],
            'AO_GUID': [],
            'AO_LEVEL': [],
            'AO_HOUSENUM': [],
            'AO_FORMALNAME': [],
            'AO_SHORTNAME': []
        })
        memory_usage = deep_getsizeof(df_main) + deep_getsizeof(df_aos) + deep_getsizeof(df_houses)
        temp_dicts = []
        add_step = 0
        added_usage = 0
        oper_time = time.time()
        for event, elem in ET.iterparse(archived_main):
            if elem.tag == 'OBJECT':
                if elem.get('ISACTUAL') == '1' and elem.get('ISACTIVE') == '1':
                    rowId = elem.get('OBJECTID')
                    objGUID = elem.get('OBJECTGUID')
                    level = int(elem.get('LEVEL'))
                    if not (rowId and objGUID and level):
                        elem.clear()
                        continue
                    if level <= 8:
                        temp_dicts.append({
                            'AO_ID' : rowId,
                            'AO_GUID': objGUID,
                            'AO_LEVEL': level,
                            'AO_FORMALNAME': elem.get('TYPENAME') + ' ' + elem.get('NAME'),
                            'AO_SHORTNAME': elem.get('TYPENAME')
                        })
                        total_log += 1
                        if total_log % 1000 == 0:
                            added_usage += deep_getsizeof(temp_dicts[add_step:])
                            memory_usage += deep_getsizeof(temp_dicts[add_step:])
                            add_step = len(temp_dicts)
                            if added_usage > 1024 * 1024 * 100:
                                df_aos = pd.concat([df_aos, pd.DataFrame(
                                    {
                                        'AO_ID': [i['AO_ID'] for i in temp_dicts],
                                        'AO_GUID': [i['AO_GUID'] for i in temp_dicts],
                                        'AO_LEVEL': [i['AO_LEVEL'] for i in temp_dicts],
                                        'AO_FORMALNAME': [i['AO_FORMALNAME'] for i in temp_dicts],
                                        'AO_SHORTNAME': [i['AO_SHORTNAME'] for i in temp_dicts]
                                    }
                                )], ignore_index=True, join='outer')
                                add_step = 0
                                added_usage = 0
                                temp_dicts.clear()
                                memory_usage = (deep_getsizeof(df_main) + deep_getsizeof(df_aos) + deep_getsizeof(
                                    df_houses) +
                                                deep_getsizeof(temp_dicts))
                            update_bars(1, 0, total_log, 'объект', 'Обработка адресных объектов...')
            elem.clear()
        if len(temp_dicts) > 0:
            df_aos = pd.concat([df_aos, pd.DataFrame(
                {
                    'AO_ID': [i['AO_ID'] for i in temp_dicts],
                    'AO_GUID': [i['AO_GUID'] for i in temp_dicts],
                    'AO_LEVEL': [i['AO_LEVEL'] for i in temp_dicts],
                    'AO_FORMALNAME': [i['AO_FORMALNAME'] for i in temp_dicts],
                    'AO_SHORTNAME': [i['AO_SHORTNAME'] for i in temp_dicts]
                }
            )], ignore_index=True, join='outer')
        else:
            if len(df_aos['AO_ID'].to_list()) == 0:
                raise RuntimeError('Подходящие адресные объекты отсутствуют')
        logging.log(logging.INFO,
                    'Обработка ' + str(len(df_aos['AO_ID'].to_list())) +
                    ' объектов заняла ' + return_time_as_str(time.time() - oper_time))
        temp_dicts.clear()
        del temp_dicts
        memory_usage = deep_getsizeof(df_main) + deep_getsizeof(df_aos) + deep_getsizeof(df_houses)
        logging.log(logging.INFO,
                    'Потребление памяти составило ' + return_memory_usage_as_str())
        progress_log = 0
        total_log = 0
        if 10 in read_levels:
            update_bars(1, 12, 0, 'объект', 'Обработка объектов домов...')
            add_step = 0
            added_usage = 0
            temp_dicts = []
            oper_time = time.time()
            for event, elem in ET.iterparse(archived_houses):
                if elem.tag == 'HOUSE':
                    if elem.get('ISACTUAL') == '1' and elem.get('ISACTIVE') == '1':
                        rowId = elem.get('OBJECTID')
                        objGUID = elem.get('OBJECTGUID')
                        if not (rowId and objGUID):
                            elem.clear()
                            continue
                        total_log += 1
                        houseType = elem.get('HOUSETYPE')
                        houseNum = elem.get('HOUSENUM')
                        if not houseType:
                            houseType = ''
                        else:
                            houseType = DICT_HOUSETYPES[houseType] + ' '
                        if houseNum:
                            temp_dicts.append({
                                'AO_ID': rowId,
                                'AO_GUID': objGUID,
                                'AO_LEVEL': 10,
                                'AO_HOUSENUM': houseNum,
                                'AO_SHORTNAME': houseType,
                                'AO_FORMALNAME': houseType + houseNum,
                            })
                        else:
                            temp_dicts.append({
                                'AO_ID': rowId,
                                'AO_GUID': objGUID,
                                'AO_LEVEL': 10,
                                'AO_HOUSENUM': -1,
                                'AO_SHORTNAME': houseType,
                                'AO_FORMALNAME': houseType,
                            })
                        if total_log % 5000 == 0:
                            t = deep_getsizeof(temp_dicts[add_step:])
                            added_usage += t
                            memory_usage += t
                            add_step = len(temp_dicts)
                            if added_usage > 1024 * 1024 * 100:
                                df_houses = pd.concat([df_houses, pd.DataFrame(
                                    {
                                        'AO_ID': [i['AO_ID'] for i in temp_dicts],
                                        'AO_GUID': [i['AO_GUID'] for i in temp_dicts],
                                        'AO_LEVEL': [i['AO_LEVEL'] for i in temp_dicts],
                                        'AO_HOUSENUM': [i['AO_HOUSENUM'] for i in temp_dicts],
                                        'AO_FORMALNAME': [i['AO_FORMALNAME'] for i in temp_dicts],
                                        'AO_SHORTNAME': [i['AO_SHORTNAME'] for i in temp_dicts]
                                    }
                                )], ignore_index=True, join='outer')
                                add_step = 0
                                added_usage = 0
                                temp_dicts.clear()
                                memory_usage = (
                                            deep_getsizeof(df_main) + deep_getsizeof(df_aos) + deep_getsizeof(df_houses) +
                                            deep_getsizeof(temp_dicts))
                            update_bars(1, 12, total_log, 'объект', 'Обработка объектов домов...')
                elem.clear()
            if len(temp_dicts) > 0:
                df_houses = pd.concat([df_houses, pd.DataFrame(
                    {
                        'AO_ID': [i['AO_ID'] for i in temp_dicts],
                        'AO_GUID': [i['AO_GUID'] for i in temp_dicts],
                        'AO_LEVEL': [i['AO_LEVEL'] for i in temp_dicts],
                        'AO_HOUSENUM': [i['AO_HOUSENUM'] for i in temp_dicts],
                        'AO_FORMALNAME': [i['AO_FORMALNAME'] for i in temp_dicts],
                        'AO_SHORTNAME': [i['AO_SHORTNAME'] for i in temp_dicts]
                    }
                )], ignore_index=True, join='outer')
            else:
                if len(df_houses['AO_ID'].to_list()) == 0:
                    raise RuntimeError('Подходящие объекты домов отсутствуют')
            logging.log(logging.INFO,
                        'Обработка ' + str(len(df_houses['AO_ID'].to_list())) +
                        ' объектов заняла ' + return_time_as_str(time.time() - oper_time))
            temp_dicts.clear()
            del temp_dicts
            df_main = pd.concat([df_main, df_aos, df_houses], ignore_index=True, join='outer')
            del df_aos
            del df_houses
            memory_usage = deep_getsizeof(df_main)
            logging.log(logging.INFO,
                        'Потребление памяти составило ' + return_memory_usage_as_str())
            progress_log = 0
        else:
            df_main = pd.concat([df_main, df_aos], ignore_index=True, join='outer')
            del df_aos
        update_bars(2, 0, 0, 'параметр', 'Обработка параметров адресных объектов...')
        total_objects = len(df_main['AO_ID'].to_list())
        list_ids = [str(o) + '=' + str(k) for k, o in enumerate(df_main['AO_ID'].to_list())]
        list_ids.sort(key=lambda x: int(x.split('=')[0]))
        list_ids_size = deep_getsizeof(list_ids)
        total_log = 0
        step_log = int(total_objects / 30)
        oper_time = time.time()
        for event, elem in ET.iterparse(archived_params):
            if elem.tag == 'PARAM':
                if elem.get('CHANGEIDEND') == '0':
                    param_id = elem.get('TYPEID')
                    param_value = elem.get('VALUE')
                    param_obj = elem.get('OBJECTID')
                    obj_index = binary_found_in_obj_list(list_ids, int(param_obj))
                    if obj_index == None:
                        elem.clear()
                        continue
                    try:
                        df_main.iloc[obj_index, DICT_PARAMS[param_id]] = str(param_value)
                        total_log += 1
                        if total_log % step_log == 0:
                            percent = (total_log / total_objects) * 100 / 9
                            if total_log % (step_log * 3) == 0:
                                memory_usage = deep_getsizeof(df_main) + list_ids_size
                            update_bars(2 + (percent * 25 / 100), percent, total_log, 'параметр',
                                        'Обработка параметров адресных объектов...')
                    except Exception:
                        elem.clear()
                        continue
            elem.clear()
        memory_usage = deep_getsizeof(df_main) + list_ids_size
        logging.log(logging.INFO,
                    'Обработка ' + str(total_log) + ' параметров адресных объектов заняла ' +
                    return_time_as_str(time.time() - oper_time))
        logging.log(logging.INFO,
                    'Скорость обработки составила ' + str(int(total_log /
                                                          (time.time() - oper_time))) + ' параметров в секунду')
        logging.log(logging.INFO,
                    'Потребление памяти составило ' + return_memory_usage_as_str())
        house_params = 0
        current = 1
        if 10 in read_levels:
            oper_time = time.time()
            object_params = total_log
            total_log = 0
            progress_log = 0
            step_log = int(total_objects / 25)
            writen = 0
            global_temp_file = open('temp_part' + str(current), 'w')
            for event, elem in ET.iterparse(archived_houses_params):
                if elem.tag == 'PARAM':
                    if elem.get('CHANGEIDEND') == '0':
                        param_id = elem.get('TYPEID')
                        param_value = elem.get('VALUE')
                        param_obj = elem.get('OBJECTID')
                        try:
                            obj_index = binary_found_in_obj_list(list_ids, int(param_obj))
                            if obj_index == None:
                                elem.clear()
                                continue
                            df_main.iloc[obj_index, DICT_PARAMS[param_id]] = str(param_value)
                            total_log += 1
                            if not pd.isna(df_main.iloc[obj_index, 17]):
                                check = 1
                                if not pd.isna(df_main.iloc[obj_index, 8]):
                                    check += 1
                                    if not pd.isna(df_main.iloc[obj_index, 9]):
                                        check += 1
                                        if not pd.isna(df_main.iloc[obj_index, 10]):
                                            check += 1
                                            if not pd.isna(df_main.iloc[obj_index, 11]):
                                                check += 1
                                                if not pd.isna(df_main.iloc[obj_index, 7]):
                                                    check += 1
                                    if check == 6:
                                        str_to_write = str(param_obj)
                                        df_main.iloc[obj_index, 21] = current
                                        for o in [7, 8, 9, 10, 11, 17]:
                                            str_to_write += '\t' + str(o) + '\t' + str(df_main.iloc[obj_index, o])
                                            df_main.iloc[obj_index, o] = np.nan
                                        str_to_write += '\n'
                                        global_temp_file.write(str_to_write)
                                        house_params += 6
                                        writen += 1
                                        if writen > step_size:
                                            global_temp_file.close()
                                            current += 1
                                            writen = 0
                                            global_temp_file = open('temp_part' + str(current), 'w')
                            if total_log % step_log == 0:
                                percent = ((object_params / total_objects) * 100 / 9 +
                                           (total_log / total_objects) * 100 / 11)
                                if total_log % (step_log * 9) == 0:
                                    memory_usage = deep_getsizeof(df_main) + list_ids_size
                                update_bars(2 + (percent * 25 / 100), percent, total_log, 'параметр',
                                            'Обработка параметров домов...')
                        except Exception:
                            elem.clear()
                            continue
                elem.clear()
            global_temp_file.close()
            current += 1
            global_temp_file = open('temp_part' + str(current), 'w')
            writen = 0
            ao_lst = df_main[df_main.AO_LEVEL == 10]['AO_ID'].to_list()
            for i in ao_lst:
                total_log += 1
                str_to_write = str(i)
                lst = []
                obj_index = binary_found_in_obj_list(list_ids, int(i))
                for o in [7, 8, 9, 10, 11, 17]:
                    test = df_main.iloc[obj_index, o]
                    if not pd.isna(test):
                        lst.append(o)
                if len(lst) > 0:
                    for o in lst:
                        str_to_write += '\t' + str(o) + '\t' + str(df_main.iloc[obj_index, o])
                        house_params += 1
                        df_main.iloc[obj_index, o] = np.nan
                    str_to_write += '\n'
                    df_main.iloc[obj_index, 21] = current
                    global_temp_file.write(str_to_write)
                    writen += 1
                    if writen > step_size:
                        global_temp_file.close()
                        current += 1
                        writen = 0
                        global_temp_file = open('temp_part' + str(current), 'w')
                if total_log % step_log == 0:
                    percent = ((object_params / total_objects) * 100 / 9 +
                               (total_log / total_objects) * 100 / 11)
                    if total_log % (step_log * 9) == 0:
                        memory_usage = deep_getsizeof(df_main) + list_ids_size
                    update_bars(2 + (percent * 25 / 100), percent, total_log, 'параметр',
                                'Обработка параметров домов...')
            global_temp_file.close()
            memory_usage = deep_getsizeof(df_main) + list_ids_size
            logging.log(logging.INFO,
                        'Обработка ' + str(total_log) +
                        ' параметров домов заняла ' + return_time_as_str(time.time() - oper_time))
            logging.log(logging.INFO,
                        'Скорость обработки составила ' + str(int(total_log /
                                                                  (time.time() - oper_time))) + ' параметров в секунду')
            logging.log(logging.INFO,
                        'Потребление памяти составило ' + return_memory_usage_as_str())
        progress_log = 0
        total_log = 0
        update_bars(27, 0, 0, 'параметр', 'Заполнение официальных наименований...')
        step_log = int(total_objects / 100)
        oper_time = time.time()
        for i in list_ids:
            now = int(i.split('=')[1])
            if pd.isna(df_main.iloc[now, 15]):
                df_main.iloc[now, 15] = df_main.iloc[now, 14]
            if df_main.iloc[now, 3] > 1:
                df_main.iloc[now, 0] = df_main.iloc[now, 1] + ': ' + df_main.iloc[now, 15]
            else:
                df_main.iloc[now, 0] = df_main.iloc[now, 15]
            total_log += 1
            if total_log % (step_log * 3) == 0:
                memory_usage = deep_getsizeof(df_main) + list_ids_size
            if total_log % step_log == 0:
                percent = (total_log / total_objects) * 100
                update_bars(27 + (percent / 100), percent, total_log, 'параметр',
                            'Заполнение официальных наименований...')
        memory_usage = deep_getsizeof(df_main) + list_ids_size
        logging.log(logging.INFO,
                    'Обработка ' + str(total_log) +
                    ' официальных наименований заняла ' + return_time_as_str(time.time() - oper_time))
        logging.log(logging.INFO,
                    'Скорость обработки составила ' + str(int(total_log /
                                                              (time.time() - oper_time))) + ' параметров в секунду')
        logging.log(logging.INFO,
                    'Потребление памяти составило ' + return_memory_usage_as_str())
        progress_log = 0
        total_log = 0
        update_bars(28, 0, 0, 'адрес', 'Заполнение муниципальных адресов...')
        del now
        step_log = int(total_objects / 300)
        oper_time = time.time()
        for event, elem in ET.iterparse(archived_mun_hierarchy):
            if elem.tag == 'ITEM':
                if elem.get('NEXTID') == '0' and elem.get('ISACTIVE') == '1':
                    objId = elem.get('OBJECTID')
                    obj_index = binary_found_in_obj_list(list_ids, int(objId))
                    if obj_index == None:
                        elem.clear()
                        continue
                    path = elem.get('PATH').split('.')
                    address = ''
                    for i in path:
                        try:
                            i_index = binary_found_in_obj_list(list_ids, int(i))
                            fragment = df_main.iloc[i_index, 15]
                            address = address + fragment + ', '
                        except Exception:
                            continue
                    address = address[:len(address) - 2]
                    df_main.iloc[obj_index, 19] = address
                    total_log += 1
                    if total_log % step_log == 0:
                        percent = (total_log / total_objects) * 10000
                        if total_log % (step_log * 10) == 0:
                            memory_usage = deep_getsizeof(df_main) + list_ids_size
                        update_bars(28 + (percent * 22 / 10000), percent**0.5, total_log, 'адрес',
                                    'Заполнение муниципальных адресов...')
            elem.clear()
        logging.log(logging.INFO,
                    'Обработка ' + str(total_log) +
                    ' муниципальных адресов заняла ' + return_time_as_str(time.time() - oper_time))
        logging.log(logging.INFO,
                    'Скорость обработки составила ' + str(int(total_log /
                                                              (time.time() - oper_time))) + ' адресов в секунду')
        logging.log(logging.INFO,
                    'Потребление памяти составило ' + return_memory_usage_as_str())
        progress_log = 0
        update_bars(50, 0, 0, 'адрес', 'Построение дерева связей по адм. иерархии...')
        total_log = 0
        oper_time = time.time()
        for event, elem in ET.iterparse(archived_hierarchy):
            if elem.tag == 'ITEM':
                if elem.get('NEXTID') == '0' and elem.get('ISACTIVE') == '1':
                    objId = elem.get('OBJECTID')
                    obj_index = binary_found_in_obj_list(list_ids, int(objId))
                    if obj_index == None:
                        elem.clear()
                        continue
                    total_log += 1
                    if total_log % step_log == 0:
                        percent = (total_log / total_objects) * 10000
                        if total_log % (step_log * 10) == 0:
                            memory_usage = deep_getsizeof(df_main) + list_ids_size
                        update_bars(50 + (percent * 35 / 10000), percent**0.5, total_log, 'адрес',
                                    'Построение дерева связей по адм. иерархии...')
                    #Выделение id из элемента PATH
                    path = elem.get('PATH').split('.')
                    address = ''
                    parent_list = []
                    for i in path:
                        try:
                            i_index = binary_found_in_obj_list(list_ids, int(i))
                            fragment = df_main.iloc[i_index, 15]
                            address = address + fragment + ', '
                            parent_list.append((i_index, df_main.iloc[i_index, 3], df_main.iloc[i_index, 16],
                                                df_main.iloc[i_index, 15]))
                        except Exception:
                            continue
                    address = address[:len(address) - 2]
                    df_main.iloc[obj_index, 18] = address
                    level = df_main.iloc[obj_index, 3]
                    self_name = df_main.iloc[obj_index, 15]
                    parent = None
                    for i in parent_list:
                        if i[1] >= level or i[3] == self_name:
                            continue
                        if i[1] in read_levels:
                            parent = list(i)
                    if parent == None:
                        continue
                    if level != 1:
                        df_main.iloc[parent[0], 20] = True
                        df_main.iloc[obj_index, 4] = df_main.iloc[parent[0], 0]
                    df_main.iloc[obj_index, 20] = True
            elem.clear()
        update_bars(85, 0, 0, 'адрес', 'Удаление избыточной информации...')
        logging.log(logging.INFO,
                    'Обработка ' + str(total_log) +
                    ' административных адресов заняла ' + return_time_as_str(time.time() - oper_time))
        logging.log(logging.INFO,
                    'Скорость обработки составила ' + str(int(total_log /
                                                              (time.time() - oper_time))) + ' адресов в секунду')
        logging.log(logging.INFO,
                    'Потребление памяти составило ' + return_memory_usage_as_str())
        oper_time = time.time()
        df_main = df_main[df_main.DETAIL == True]
        df_main = df_main.astype(DICT_DF_DTYPES)
        list_ids.clear()
        del list_ids
        memory_usage_old = memory_usage
        memory_usage = memory_usage_old - deep_getsizeof(df_main)
        logging.log(logging.INFO,
                    'Чистка лишних объектов заняла ' + return_time_as_str(time.time() - oper_time))
        logging.log(logging.INFO,
                    'Памяти освобождено: ' + return_memory_usage_as_str())
        memory_usage = deep_getsizeof(df_main)
        #Запись эксель-таблиц
        oper_time = time.time()
        progress_log = 0
        total_log = 0
        for i in read_levels:
            if i != 10:
                update_bars(86, 100 * total_log / len(read_levels), total_log, 'файл',
                            'Запись эксель-таблиц адресных объектов...')
                writer = pd.ExcelWriter('Data/' + str(n) + '/' + DICT_LEVELS[i] + '.xlsx', engine='xlsxwriter')
                df_main[df_main.AO_LEVEL == i].drop(columns=['DETAIL', 'FILE']).to_excel(writer, 'Лист1', index=False)
                writer.close()
            total_log += 1
        update_bars(86, 100, total_log, 'файл', 'Запись таблицы улиц завершена...')
        del writer
        if 10 in read_levels:
            df_main = df_main[df_main.AO_LEVEL == 10]
            memory_usage_old = memory_usage
            memory_usage = memory_usage_old - deep_getsizeof(df_main)
            logging.log(logging.INFO,
                        'Построение таблиц адресных объектов заняло ' + return_time_as_str(time.time() - oper_time))
            logging.log(logging.INFO,
                        'После выполнения операции было освобождено памяти: ' + return_memory_usage_as_str())
            progress_log = 0
            update_bars(87, 0, 0, 'параметр', 'Построение таблиц для объектов домов...')
            memory_usage = deep_getsizeof(df_main)
            total_log = 0
            step_log = int(house_params / 500)
            percent = 0
            oper_time = time.time()
            for i in range(1, current + 1):
                df_temp = df_main[df_main.FILE == i].drop(columns=['DETAIL', 'FILE'])
                df_main = df_main[df_main.FILE != i]
                global_temp_file = open('temp_part' + str(i), 'r')
                list_ids = [[str(o) + '=' + str(k) for k, o in enumerate(df_temp['AO_ID'].to_list()) if str(o)[0] == str(m)]
                            for m in range(1, 10)]
                for m in range(0, 9):
                    list_ids[m].sort(key=lambda x: int(x.split('=')[0]))
                for line in global_temp_file:
                    splitted = line.split('\t')
                    ind = int(splitted[0][0]) - 1
                    obj_index = binary_found_in_obj_list(list_ids[ind], int(splitted[0]))
                    if obj_index == None:
                        continue
                    for q in range(0, int((len(splitted) - 1) / 2)):
                        df_temp.iloc[obj_index, int(splitted[1 + q * 2])] = splitted[2 + q * 2]
                        total_log += 1
                    if total_log % step_log == 0:
                        percent = (total_log / (house_params + 1)) * 100
                        update_bars(87 + (percent * 13 / 100), percent, total_log, 'параметр',
                                    'Построение таблиц для объектов домов...')
                global_temp_file.close()
                list_ids.clear()
                os.remove('temp_part' + str(i))
                writer = pd.ExcelWriter('Data/' + str(n) + '/' + DICT_LEVELS[10] + '_part' + str(i) + '.xlsx',
                                        engine='xlsxwriter')
                df_temp.to_excel(writer, 'Лист1', index=False)
                writer.close()
                del df_temp
                memory_usage = deep_getsizeof(df_main) + deep_getsizeof(list_ids)
                progress_log = 0
                update_bars(87 + (percent * 13 / 100), percent, 0, 'параметр',
                            'Построение таблиц для объектов домов...')
            del writer
            memory_usage = 0
            progress_log = 0
            update_bars(100, 100, 0, 'параметр', 'Обработка региона завершена!')
            logging.log(logging.INFO,
                        'Построение таблиц домов заняло ' + return_time_as_str(time.time() - oper_time))
            logging.log(logging.INFO,
                        'Скорость обработки составила ' + str(int(total_log /
                                                                  (time.time() - oper_time))) +
                        ' параметров домов в секунду')
        logging.log(logging.INFO,
                    'Всего для обработки данного региона понадобилось ' + return_time_as_str(time.time() - subj_time))


def main():
    global subject_num, label_value, label_subj_value, progress_main, progress_add, pb_main, pb_add, root, label
    global label_subj, read_levels, houses_per_excel
    #Чтение файла с параметрами properties
    properties = open('properties')
    subj_params = ['ALL']
    for line in properties:
        name = line.split(': ')[0]
        if name == "ReadSubjects":
            subj_params = line.split(': ')[1][1:len(line.split(': ')[1]) - 2].split(', ')
        if name == "ReadLevels":
            temp_params = line.split(': ')[1][1:len(line.split(': ')[1]) - 2].split(', ')
            if len(temp_params) > 0:
                read_levels = [int(i) for i in temp_params]
        if name == "HousesPerExcel":
            houses_per_excel = int(line.split(': ')[1])
            break
    properties.close()
    del properties
    del temp_params

    #Подготовка списка регионов: проверка, были ли они уже обработаны и какие из них необходимо обработать
    if not os.path.exists('Data'):
        os.mkdir('Data')
    if subj_params[0] == 'ALL':
        subj_list = [i for i in range(0, 100) if not os.path.exists('Data\\' + str(i))]
    else:
        subj_list = [i for i in range(0, 100) if not os.path.exists('Data\\' + str(i)) and str(i) in subj_params]

    #Главный цикл: обработка регионов
    for i in subj_list:
        #Инициализация логов
        if not os.path.exists('log'):
            os.mkdir('log')
        file_log = open('log/' + str(i) + '.log', mode='w')
        file_log.close()
        logging.basicConfig(filename=('log/' + str(i) + '.log'), level=logging.INFO, force=True)

        #Подготовка региона и переменных прогресса его обработки
        subject_num = i
        label_value = 'Подготовка...'
        label_subj_value = str('В обработке субъект №' + str(subject_num))
        progress_main = 0
        progress_add = 0
        pb_main['value'] = progress_main
        pb_add['value'] = progress_add
        label['text'] = label_value
        label_subj['text'] = label_subj_value
        root.update()

        #Вызов функции. В случае ошибки - её логирование
        try:
            read_subject(i)
        except Exception as err:
            logging.error(time.ctime(time.time()) + ': ' + str(err))


if __name__ == '__main__':
    main()

