# -*- coding: utf-8 -*-
import clr
import sys
import operator
import os
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('RevitServices')
clr.AddReference('RevitAPI')
import RevitServices
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import FilteredElementCollector as FEC
from Autodesk.Revit.UI import UIApplication
from Autodesk.Revit import UI
from RevitServices.Persistence import DocumentManager
clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)
from RevitServices.Transactions import TransactionManager


import pyrevit
from pyrevit import forms
from pyrevit import HOST_APP
from pyrevit import revit

from StructureUpdater import types_updater
from TypeParametersUpdater import parameters_updater

uiapp = __revit__ # noqa
app = __revit__.Application # noqa
uiapp = HOST_APP.uiapp
# uidoc = __revit__.ActiveUIDocument # noqa
doc = __revit__.ActiveUIDocument.Document # noqa
app_version = __revit__.Application.VersionNumber # noqa

from datetime import datetime

"""Read and Write Excel Files."""
#pylint: disable=import-error
import xlrd
import xlsxwriter

def find_cell_coordinates(excel_path, target_value):
    # Open the Excel workbook
    workbook = xlrd.open_workbook(excel_path)
    # Iterate through all sheets
    row_ind = []
    col_ind = []
    sheet_index = []
    for sheet_index_ in range(workbook.nsheets):
        sheet = workbook.sheet_by_index(sheet_index_)
        # Iterate through all rows and columns
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                cell_value = str(sheet.cell_value(row, col))
                mod_cell_value = cell_value
                
                # Check if the cell contains the target value
                if mod_cell_value == str(target_value):
                    row_ind.append(row + 1)
                    col_ind.append(col + 1)
                    sheet_index.append(sheet_index_)
    if isinstance(row_ind, list) and len(row_ind)>1:
        return row_ind, col_ind, sheet_index
    else:
        try:
            row_ind = row_ind[0]
            col_ind = col_ind[0]
            sheet_index = sheet_index[0]
        except:
            pass
        return row_ind, col_ind, sheet_index
    

def find_cell_coordinates_on_sheet(excel_path, sheet_index, target_value):
    # Open the Excel workbook
    workbook = xlrd.open_workbook(excel_path)
    # Iterate through all sheets
    row_ind = []
    col_ind = []
    sheet = workbook.sheet_by_index(sheet_index)
    # Iterate through all rows and columns
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            cell_value = str(sheet.cell_value(row, col))
            mod_cell_value = cell_value
            
            # Check if the cell contains the target value
            if mod_cell_value == str(target_value):
                row_ind.append(row + 1)
                col_ind.append(col + 1)
    if isinstance(row_ind, list) and len(row_ind)>1:
        return row_ind, col_ind
    else:
        try:
            row_ind = row_ind[0]
            col_ind = col_ind[0]
        except:
            pass
        return row_ind, col_ind

def find_value_in_excel(excel_path, target_value):
    # Open the Excel workbook
    workbook = xlrd.open_workbook(excel_path)

    # Iterate through all sheets
    for sheet_index in range(workbook.nsheets):
        sheet = workbook.sheet_by_index(sheet_index)

        # Iterate through all rows and columns
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                cell_value = sheet.cell_value(row, col)
                
                # Check if the cell contains the target value
                if cell_value == target_value:
                    return cell_value
                
def find_all_indexes(main_string, substring):
    indexes = []
    start_index = 0

    while True:
        index = main_string.find(substring, start_index)
        if index == -1:
            break
        indexes.append(index)
        start_index = index + 1

    return indexes

def flatten(element, flat_list=None):
    if flat_list is None:
        flat_list = []
    if hasattr(element, "__iter__"):
        for item in element:
            flatten(item, flat_list)
    else:
        flat_list.append(element) 
    return flat_list


def create_param_dictionary(excel_path, target_sheet, sheet_index, row_ind, parameter_list_toextract):
    param_dict = {}
    # sheet_ind_list = find_cell_coordinates(excel_path, 'CP_Mat_Finish_01')[2] # Определяем индексы всех листов, где есть параметр материала
    # print('row_ind -' + '{}'.format(row_ind))
    # print('check00_mark_row_list')

    for param in parameter_list_toextract:
        # print(param)
        param_rows_list = find_cell_coordinates_on_sheet(excel_path, sheet_index, param)[0] #извлекаем все строки параметра материала на данном листе
        # print param_rows_list
        param_found = find_cell_coordinates_on_sheet(excel_path, sheet_index, param)[1]  # извлекаем индекс столбца параметра материала на данном листе
        if param_found:
            # print('CHECK3')
            if isinstance(param_rows_list, list) and len(param_rows_list)>1:
                # print('CHECK4')
                param_col_ind = find_cell_coordinates_on_sheet(excel_path, sheet_index, param)[1][0]-1  # извлекаем индекс столбца параметра материала на данном листе
                # print('check1')
                if row_ind in param_rows_list:
                    key_param = param
                    material_description = target_sheet.cell_value(row_ind-1, param_col_ind+2)
                    thickness = target_sheet.cell_value(row_ind-1, param_col_ind+1)
                    material_name = target_sheet.cell_value(row_ind-1, param_col_ind+3)
                    if len(material_description)>0:
                        # print(param_dict)
                        # print('CHECK5')
                        if param not in param_dict.keys():
                            param_dict[param] = []
                        if not thickness>0:
                            print('!!! ТОЛЩИНА МАТЕРИАЛА {} НЕ ОПРЕДЕЛЕНА, ОТМЕНА ОПЕРАЦИИ !!!'.format(param))
                            return None
                        elif not len(material_name)>0:
                            print('!!! ИМЯ МАТЕРИАЛА {} НЕ ОПРЕДЕЛЕНО, ОТМЕНА ОПЕРАЦИИ !!!'.format(param))
                            return None
                        param_dict[param].append(thickness)
                        # print('KEY {} - VALUE {}'.format(param, thickness))
                        param_dict[param].append(material_description)
                        param_dict[param].append(material_name)
                    elif param == 'FLOOR_ATTR_THICKNESS_PARAM':
                        # print('CHECK6')
                        if thickness>0:
                            param_dict[param] = thickness
                        else:
                            print('!!! ТОЛЩИНА ПИРОГА {} НЕ ОПРЕДЕЛЕНА, ОТМЕНА ОПЕРАЦИИ !!!'.format(param))
                            return None
                    if param == 'CP_Mat_Finish_01':
                        # print('CHECK7')
                        type_comments = target_sheet.cell_value(row_ind-1, param_col_ind-1)
                        function = target_sheet.cell_value(row_ind-1, param_col_ind-3)
                        if not len(type_comments)>0:
                            print('!!! НАИМЕНОВАНИЕ ТИПА НЕ ОПРЕДЕЛЕНО, ОТМЕНА ОПЕРАЦИИ !!!')
                            return None
                        elif not len(function)>0:
                            print('!!! ФУНКЦИЯ ТИПА НЕ ОПРЕДЕЛЕНА, ОТМЕНА ОПЕРАЦИИ !!!')
                            return None
                        param_dict[param].append(type_comments)
                        param_dict[param].append(function)
                            # Parameter_List_Internal.remove(param)
                            # exec("print '{} - ' + '{}'".format(param, material_description))
            else:
                # print(excel_path)
                # print(sheet_index)
                # print(param)
                # print(find_cell_coordinates_on_sheet(excel_path, sheet_index, param)[1])
                param_col_ind = find_cell_coordinates_on_sheet(excel_path, sheet_index, param)[1]-1
                # print('check3')
                if row_ind == param_rows_list:
                    key_param = param
                    # print('check04')
                    material_description = target_sheet.cell_value(row_ind-1, param_col_ind+2)
                    thickness = target_sheet.cell_value(row_ind-1, param_col_ind+1)
                    material_name = target_sheet.cell_value(row_ind-1, param_col_ind+3)
                    if not thickness>0:
                        print('!!! ТОЛЩИНА МАТЕРИАЛА {} НЕ ОПРЕДЕЛЕНА, ОТМЕНА ОПЕРАЦИИ !!!'.format(param))
                        return None
                    elif not len(material_name)>0:
                        print('!!! ИМЯ МАТЕРИАЛА {} НЕ ОПРЕДЕЛЕНО, ОТМЕНА ОПЕРАЦИИ !!!'.format(param))
                        return None
                    if len(material_description)>0:
                        if param not in param_dict.keys():
                            param_dict[param] = []
                        param_dict[param].append(thickness)
                        param_dict[param].append(material_description)
                        param_dict[param].append(material_name)
                    elif param == 'FLOOR_ATTR_THICKNESS_PARAM':
                        # print('CHECK3')
                        if thickness>0:
                            param_dict[param] = thickness
                        else:
                            print('!!! ТОЛЩИНА ПИРОГА {} НЕ ОПРЕДЕЛЕНА, ОТМЕНА ОПЕРАЦИИ !!!'.format(param))
                            return None
                    if param == 'CP_Mat_Finish_01':
                        # print('CHECK3')
                        type_comments = target_sheet.cell_value(row_ind-1, param_col_ind-1)
                        function = target_sheet.cell_value(row_ind-1, param_col_ind-3)
                        if not len(type_comments)>0:
                            print('!!! НАИМЕНОВАНИЕ ТИПА НЕ ОПРЕДЕЛЕНО, ОТМЕНА ОПЕРАЦИИ !!! ')
                            return None
                        elif not len(function)>0:
                            print('!!! ФУНКЦИЯ ТИПА НЕ ОПРЕДЕЛЕНА, ОТМЕНА ОПЕРАЦИИ !!!')
                            return None
                        param_dict[param].append(type_comments)
                        param_dict[param].append(function)
    count = 0
    dict_correct = True
    key_value = ''
    for key, v_list in param_dict.items():
        if isinstance(v_list, list):
            count = len(v_list)
        else:
            count = 1
        key_value = key
    # print(key_value)
    # print(count)

    image_param_col = find_cell_coordinates_on_sheet(excel_path, sheet_index, 'CP_Gen_Image')[1]
    image_link = target_sheet.cell_value(row_ind-1, image_param_col-1)
    if image_link:
        param_dict['CP_Gen_Image'] = image_link

    if key_value == 'CP_Mat_Finish_01':
        if count != 5:
            dict_correct = False
    elif key_value == 'FLOOR_ATTR_THICKNESS_PARAM':
        if count != 1:
            dict_correct = False
    elif 'CP_Mat_Finish_' in key_value and count != 3:
        # print('CHECK')
        dict_correct = False
    # print(dict_correct)
    if dict_correct:
        return param_dict
    return None




#USER INPUTS
excel_path = pyrevit.forms.pick_file(file_ext='xls', files_filter='', init_dir='', restore_dir=True, multi_file=False, unc_paths=False, title='Выбери файл Excel Реестра')

# Проверяем, если пользователь нажал "Отмена"
if not excel_path:
    print("Операция отменена пользователем.")
    sys.exit()  # Корректно завершаем выполнение скрипта без ошибки

workbook = xlrd.open_workbook(excel_path)

sheets = [workbook.sheet_by_index(sheet_index) for sheet_index in range(workbook.nsheets)]
sheets = sheets[1:]
selected_sheets = pyrevit.forms.SelectFromList.show(sheets, title = 'Выбери Каталог', multiselect = True, name_attr='name', button_name = 'Выбрать')

# Проверяем, если пользователь нажал "Отмена"
if not selected_sheets:
    print("Операция отменена пользователем.")
    sys.exit()  # Корректно завершаем выполнение скрипта без ошибки

marks = []
type_row_dict = {}
for sheet in selected_sheets:
    # Iterate through all rows and columns
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            cell_value = str(sheet.cell_value(row, col))
            # Check if the cell contains the target value
            if cell_value == 'CP_Gen_Mark':
                for row in range(sheet.nrows):
                    mark = str(sheet.cell_value(row, col))
                    marks.append(mark)
elements_to_exclude = ["CP_Gen_Mark", 'Марка типа', 'Марка', '']
marks = [element for element in marks if element not in elements_to_exclude]

marks_check = set(marks)
marks_not_modified = []
marks_modified = []
for mark in marks_check:
    score_ind = mark.find('_')
    # print slash_ind
    if score_ind != -1:
        marks_not_modified.append(mark[:score_ind])
    else:
        marks_modified.append(mark)

marks_to_remove = []

for m in marks_modified:
    if m in marks_not_modified:
        marks_to_remove.append(m)

for m in marks_to_remove:
    marks_modified.remove(m)

selected_marks = pyrevit.forms.SelectFromList.show(list(marks_modified), title = 'Выбери Марки', multiselect = True, button_name = 'Выбрать')

# Проверяем, если пользователь нажал "Отмена"
if not selected_marks:
    print("Операция отменена пользователем.")
    sys.exit()  # Корректно завершаем выполнение скрипта без ошибки

if len(selected_marks) == 0:
    selected_marks = marks_modified

# print selected_marks

Parameter_List_Internal = [
    'CP_Mat_Finish_01',
    'CP_Mat_Finish_02',
    'CP_Mat_Finish_03',
    'CP_Mat_Finish_04',
    'CP_Mat_Finish_05',
    'CP_Mat_Finish_06',
    'CP_Mat_Finish_07',
    'CP_Mat_Finish_08',
    'CP_Mat_Finish_09',
    'FLOOR_ATTR_THICKNESS_PARAM'
]

Categories_Dict = {
    'FL': [FloorType, 'Перекрытия'],
    'WL': [WallType, 'Стены'],
    'CL': [CeilingType, 'Потолки'],
}


#Собираем все параметры семейства в лист


fam_type_input_names = [
    "CP_Gen_Mark",
    "ALL_MODEL_TYPE_COMMENTS",
    "FLOOR_ATTR_THICKNESS_PARAM"
]

lib_ind = excel_path.find('01_БИБЛИОТЕКА_ПРОЕКТА')
library_folder_path = excel_path[:lib_ind+len('01_БИБЛИОТЕКА_ПРОЕКТА')+1] + '01_СЕМЕЙСТВА\\'

type_dict = {}
for target_value in selected_marks:
    print('\n ================== марка {} - ЧТЕНИЕ ДАННЫХ ИЗ РЕЕСТРА========================='.format(target_value))

    # param_dict = {}
    merged_param_dict = {}
    # print(target_value)
    #Извлекаем значения листа и строк типоразмеров, соотвествующих марке семейства
    # mark_row_list = []
    mark_row_list = sorted(list(set(find_cell_coordinates(excel_path, target_value)[0])))
    # print('CHECK501')
    print('строки в Excel - {}'.format(mark_row_list))
    # print('CHECK501')
    multiple_rows = False
    if isinstance(mark_row_list, list):
        multiple_rows = True
    if multiple_rows:
        sheet_index = find_cell_coordinates(excel_path, target_value)[2][0]
    else:
        sheet_index = find_cell_coordinates(excel_path, target_value)[2]
    target_sheet = workbook.sheet_by_index(sheet_index)
    print('\n Код категории в реестре - {}'.format(target_sheet.name))

    #Собираем все значения параметров из реестра в словарь
    if multiple_rows:
        for row_ind in mark_row_list:
            param_dict = create_param_dictionary(excel_path, target_sheet, sheet_index, row_ind, Parameter_List_Internal)
            # print(row_ind)
            # print(param_dict)
            if param_dict is not None:
                merged_param_dict.update(param_dict)
            # print('check01')
    else:
        param_dict = create_param_dictionary(excel_path, target_sheet, sheet_index, mark_row_list, Parameter_List_Internal)
        if param_dict is not None:
            merged_param_dict = param_dict
        # print('check02')
    
    param_count = 0
    image_key = False
    for k in merged_param_dict.keys():
        if k != 'CP_Gen_Image':
            param_count += 1
        else:
            image_key = True
    if not image_key:
        print('!!! НЕ НАЙДЕНО ИЗОБРАЖЕНИЕ В РЕЕСТРЕ !!!')
    
    
    if param_count == len(mark_row_list) and image_key:
    
        fam_category = Categories_Dict[target_sheet.name][0]
        merged_param_dict['Category'] = fam_category
        mark = target_value
        merged_param_dict['CP_Gen_Mark'] = mark
        comments = merged_param_dict['CP_Mat_Finish_01'][3]
        # print('check03')
        # print(merged_param_dict['CP_Gen_Mark'] )
        # print(merged_param_dict['FLOOR_ATTR_THICKNESS_PARAM'])
        # if multiple_rows:
        #     print('check11')
        #     thickness_total = merged_param_dict['FLOOR_ATTR_THICKNESS_PARAM'][0]
        # else:
        thickness_total = merged_param_dict['FLOOR_ATTR_THICKNESS_PARAM']
        # print('check12')
        # print(thickness_total)
        category_name = Categories_Dict[target_sheet.name][1]
        if "ПП" in mark:
            prefix_1 = 'FL_'
        elif 'ПО' in mark:
            prefix_1 = 'CL_'
        elif 'П' in mark:
            prefix_1 = 'FFL_'
        elif 'СТ' in mark:
            prefix_1 = 'WL_'
        elif 'ОТ' in mark:
            prefix_1 = 'WF_'
        prefix_2 = merged_param_dict['CP_Mat_Finish_01'][4]
        # print('check02')


        #Проверка прохода по типам и параметрам
        exec("print 'Марка - {}'".format(mark))
        if prefix_1 and prefix_2:
            # type_name = 'type'
            type_name = prefix_1 + prefix_2 + '_' + mark + '_'  + comments + '_' + str(int(thickness_total)) + 'мм'
            exec("print 'Имя типоразмера - {}'".format(type_name))
        else:
            if not prefix_1:
                print('!!! ДЛЯ ТИПА  НЕ НАЙДЕНА КАТЕГОРИЯ ДЛЯ МАРКИ !!!')
            if not prefix_2:
                print('!!! ДЛЯ ТИПА НЕ УКАЗАНА ФУНКЦИЯ !!!')
        # print('check01')
        if thickness_total:
            exec("print 'Общая толщина пирога - {}' + 'мм'".format(thickness_total))
        else:
            print('!!! ДЛЯ ТИПА НЕ НАЙДЕН ПАРАМЕТР FLOOR_ATTR_THICKNESS_PARAM !!!')

        print('\n КЛЮЧЕВЫЕ ПАРАМЕТРЫ ТИПА:')
        for par, value in sorted(merged_param_dict.items()):
            if par == 'Category':
                if value:
                    print('Категория - {}'.format(category_name))
                else:
                    print('!!! КАТЕГОРИЯ ТИПА НЕ ОПРЕДЕЛЕНА !!!')
                # print("___")
            elif par == 'CP_Gen_Image':
                if value:
                    print('\n Путь к изображению - {}'.format(value))
                else:
                    print('!!! ССЫЛКА НА ИЗОБРАЖЕНИЕ НЕ НАЙДЕНА !!!')
                # print("___")
            elif par == 'CP_Gen_Mark':
                if value:
                    print('Марка типа - {}'.format(value))
                else:
                    print('!!! МАРКА ТИПА НЕ ОПРЕДЕЛЕНА !!!')
                # print("___")
            else:
                if isinstance(value, list):
                    if len(value)>1:
                        print('\n{}:'.format(par))
                        print('Толщина - {}'.format(value[0]))
                        print('Описание материала - {}'.format(value[1]))
                        print('Материал Revit - {}'.format(value[2]))
                    if len(value)>3:
                        print('Комментарии к типоразмеру - {}'.format(value[3]))
                        print('Функция - {}'.format(value[4]))
                    # print("___")
        # print(param_count)
        # print(len(mark_row_list))
        type_dict[type_name] = merged_param_dict

updated_types = types_updater(doc, type_dict)
updated_parameters = parameters_updater(doc, type_dict)