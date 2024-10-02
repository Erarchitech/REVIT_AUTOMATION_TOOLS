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
                # cell_dot_ind = cell_value.find('.')
                # cell_score_ind = cell_value.find('_')
                # if cell_dot_ind != -1 and cell_score_ind == -1:
                #     mod_cell_value = cell_value[:cell_dot_ind]
                # else:
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
            # cell_dot_ind = cell_value.find('.')
            # cell_score_ind = cell_value.find('_')
            # if cell_dot_ind != -1 and cell_score_ind == -1:
            #     mod_cell_value = cell_value[:cell_dot_ind]
            # else:
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

# Функция для получения значения параметра толщины
def get_thickness_parameter(element):
    # Проверяем тип элемента (например, стена, перекрытие или потолок)
    if element.Category.Id.IntegerValue == int(BuiltInCategory.OST_Walls):
        param = element.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)
    elif element.Category.Id.IntegerValue == int(BuiltInCategory.OST_Floors):
        param = element.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)
    elif element.Category.Id.IntegerValue == int(BuiltInCategory.OST_Ceilings):
        param = element.get_Parameter(BuiltInParameter.CEILING_THICKNESS_PARAM)
    else:
        param = None

    # Если параметр найден и его тип данных - Double, то возвращаем значение
    if param and param.StorageType == StorageType.Double:
        # Значение параметра в футах, переводим в метры (если нужно)
        return param.AsDouble() * 0.3048  # Преобразование из футов в метры
    return None

def create_param_dictionary(excel_path, target_sheet, sheet_index, row_ind, parameter_list_toextract):
    # sheet_ind_list = find_cell_coordinates(excel_path, 'CP_Mat_Finish_01')[2] # Определяем индексы всех листов, где есть параметр материала
    # print('row_ind -' + '{}'.format(row_ind))
    # print('check00_mark_row_list')
    image_param_col = find_cell_coordinates_on_sheet(excel_path, sheet_index, 'CP_Gen_Image')[1]
    print('image_param_col - ' + str(image_param_col))
    print(row_ind)
    print(target_sheet)
    print(target_sheet.name)
    image_link = target_sheet.cell_value(row_ind-1, image_param_col-1)
    print('CHECK')
    print(len(image_link))
    print(str(image_link))
    if image_link:
        param_dict['CP_Gen_Image'] = image_link
        print('CHECK2')

    for param in parameter_list_toextract:
        param_rows_list = find_cell_coordinates_on_sheet(excel_path, sheet_index, param)[0] #извлекаем все строки параметра материала на данном листе
        # print param_rows_list
        param_found = find_cell_coordinates_on_sheet(excel_path, sheet_index, param)[1]  # извлекаем индекс столбца параметра материала на данном листе
        if param_found:
            if isinstance(param_rows_list, list) and len(param_rows_list)>1:
                param_col_ind = find_cell_coordinates_on_sheet(excel_path, sheet_index, param)[1][0]-1  # извлекаем индекс столбца параметра материала на данном листе
                # print('check1')
                if row_ind in param_rows_list:
                    # print('check2')
                    material_description = target_sheet.cell_value(row_ind-1, param_col_ind+2)
                    thickness = target_sheet.cell_value(row_ind-1, param_col_ind+1)
                    material_name = target_sheet.cell_value(row_ind-1, param_col_ind+3)
                    if len(material_description)>0:
                        if param not in param_dict.keys():
                            param_dict[param] = []
                        param_dict[param].append(thickness)
                        param_dict[param].append(material_description)
                        param_dict[param].append(material_name)
                    elif param == 'FLOOR_ATTR_THICKNESS_PARAM':
                        print('check04')
                        # if param not in param_dict.keys():
                        #     param_dict[param] = []
                        # param_dict[param].append(thickness)
                        param_dict[param] = thickness
                        print('check03')
                    if param == 'CP_Mat_Finish_01':
                        type_comments = target_sheet.cell_value(row_ind-1, param_col_ind-1)
                        function = target_sheet.cell_value(row_ind-1, param_col_ind-3)
                        param_dict[param].append(type_comments)
                        param_dict[param].append(function)
                            # Parameter_List_Internal.remove(param)
                            # exec("print '{} - ' + '{}'".format(param, material_description))
            else:
                print(excel_path)
                print(sheet_index)
                print(param)
                print(find_cell_coordinates_on_sheet(excel_path, sheet_index, param)[1])
                param_col_ind = find_cell_coordinates_on_sheet(excel_path, sheet_index, param)[1]-1
                # print('check3')
                if row_ind == param_rows_list:
                    # print('check04')
                    material_description = target_sheet.cell_value(row_ind-1, param_col_ind+2)
                    thickness = target_sheet.cell_value(row_ind-1, param_col_ind+1)
                    material_name = target_sheet.cell_value(row_ind-1, param_col_ind+3)
                    if len(material_description)>0:
                        if param not in param_dict.keys():
                            param_dict[param] = []
                        param_dict[param].append(thickness)
                        param_dict[param].append(material_description)
                        param_dict[param].append(material_name)
                    elif param == 'FLOOR_ATTR_THICKNESS_PARAM':
                        print('check05')
                        # if param not in param_dict.keys():
                        #     param_dict[param] = []
                        # param_dict[param].append(thickness)
                        param_dict[param] = thickness
                        print('check06')
                    if param == 'CP_Mat_Finish_01':
                        type_comments = target_sheet.cell_value(row_ind-1, param_col_ind-1)
                        function = target_sheet.cell_value(row_ind-1, param_col_ind-3)
                        param_dict[param].append(type_comments)
                        param_dict[param].append(function)
    return param_dict


#USER INPUTS
excel_path = pyrevit.forms.pick_file(file_ext='xls', files_filter='', init_dir='', restore_dir=True, multi_file=False, unc_paths=False, title='Выбери файл Excel Реестра')
workbook = xlrd.open_workbook(excel_path)

sheets = [workbook.sheet_by_index(sheet_index) for sheet_index in range(workbook.nsheets)]
sheets = sheets[1:]
selected_sheets = pyrevit.forms.SelectFromList.show(sheets, title = 'Выбери Каталог', multiselect = True, name_attr='name', button_name = 'Выбрать')

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
    print '==================ЧТЕНИЕ ДАННЫХ ИЗ РЕЕСТРА========================='

    param_dict = {}
    merged_param_dict = {}
    
    #Извлекаем значения листа и строк типоразмеров, соотвествующих марке семейства
    # mark_row_list = []
    mark_row_list = find_cell_coordinates(excel_path, target_value)[0]
    print('CHECK501')
    print mark_row_list
    print('CHECK501')
    multiple_rows = False
    if isinstance(mark_row_list, list):
        multiple_rows = True
    if multiple_rows:
        sheet_index = find_cell_coordinates(excel_path, target_value)[2][0]
    else:
        sheet_index = find_cell_coordinates(excel_path, target_value)[2]
    target_sheet = workbook.sheet_by_index(sheet_index)
    exec("print 'Код каталога в реестре - '+ '{}'".format(target_sheet.name))

    #Собираем все значения параметров из реестра в словарь
    if multiple_rows:
        for row_ind in mark_row_list:
            param_dict = create_param_dictionary(excel_path, target_sheet, sheet_index, row_ind, Parameter_List_Internal)
            merged_param_dict.update(param_dict)
            print('check01')
    else:
        param_dict = create_param_dictionary(excel_path, target_sheet, sheet_index, mark_row_list, Parameter_List_Internal)
        merged_param_dict = param_dict
        print('check02')
    
    fam_category = Categories_Dict[target_sheet.name][0]
    merged_param_dict['Category'] = fam_category
    mark = target_value
    merged_param_dict['CP_Gen_Mark'] = mark
    comments = merged_param_dict['CP_Mat_Finish_01'][3]
    print('check03')
    print(merged_param_dict['CP_Gen_Mark'] )
    print(merged_param_dict['FLOOR_ATTR_THICKNESS_PARAM'])
    # if multiple_rows:
    #     print('check11')
    #     thickness_total = merged_param_dict['FLOOR_ATTR_THICKNESS_PARAM'][0]
    # else:
    thickness_total = merged_param_dict['FLOOR_ATTR_THICKNESS_PARAM']
    print('check12')
    print(thickness_total)
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
    if len(prefix_1)>0 and len(prefix_2)>0:
        # type_name = 'type'
        type_name = prefix_1 + prefix_2 + '_' + mark + '_'  + comments + '_' + str(int(thickness_total)) + 'мм'
        exec("print 'Имя типоразмера - {}'".format(type_name))
    else:
        print('В реестре не указана функция')
    print('check01')
    exec("print 'Общая толщина пирога - {}' + 'мм'".format(thickness_total))

    
    for par, value in merged_param_dict.items():
        if par == 'Category':
            exec("print 'Категория - ' + '{}'".format(category_name))
            print("___")
        elif par == 'CP_Gen_Image':
            exec("print 'Путь к изображению - ' + '{}'".format(value))
            print("___")
        elif par == 'CP_Gen_Mark':
            exec("print 'Марка типа - ' + '{}'".format(value))
            print("___")
        else:
            if isinstance(value, list):
                if len(value)>1:
                    exec("print '{}'".format(par))
                    exec("print 'Толщина - ' + '{}'".format(value[0]))
                    exec("print 'Описание материала - ' + '{}'".format(value[1]))
                    exec("print 'Материал Revit - ' + '{}'".format(value[2]))
                if len(value)>3:
                    exec("print 'Комментарии к типоразмеру - ' + '{}'".format(value[3]))
                    exec("print 'Функция - ' + '{}'".format(value[4]))
                print("___")


    type_dict[type_name] = merged_param_dict

updated_types = types_updater(doc, type_dict)
updated_parameters = parameters_updater(doc, type_dict)