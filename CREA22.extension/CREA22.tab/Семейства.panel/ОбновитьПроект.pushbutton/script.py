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
                cell_dot_ind = cell_value.find('.')
                cell_score_ind = cell_value.find('_')
                if cell_dot_ind != -1 and cell_score_ind == -1:
                    mod_cell_value = cell_value[:cell_dot_ind]
                else:
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

# Function to iterate through files in a folder and its subfolders
def iterate_files(folder, family_names):
    doc_paths = []
    for family_name in family_names:
        for root, dirs, files in os.walk(folder):
            for file in files:
                if file.endswith(".rfa") and family_name == file:
                    doc_paths.append(os.path.join(root, file))
    return doc_paths

def close_inactive_docs(uiapp, save_modified=False, close_ui_docs=False):
    for doc in uiapp.Application.Documents:
        if not close_ui_docs and UI.UIDocument(doc).GetOpenUIViews():
            continue
        name = doc.Title
        doc.Close(save_modified)
        print 'Closed: {}'.format(name)


# Function to open and close family documents

class FamilyOption(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues = True
        return True
    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        overwriteParameterValues = False
        source = FamilySource.Family
        return True
    
def family_reload(filePath, loadOptions, ProjectDoc = doc):
    try:
        ProjectDoc.LoadFamily(filePath, loadOptions)
        return 1
    except:
        return 0


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
selected_marks = pyrevit.forms.SelectFromList.show(marks, title = 'Выбери Марки', multiselect = True, button_name = 'Выбрать')

if len(selected_marks) == 0:
    selected_marks = marks

# print selected_marks





Parameter_List_Internal = [
    "CP_Gen_Mark",
    "CP_Dim_Length",
    "CP_Dim_Width",
    "CP_Dim_Height",
    "CP_Gen_Description",
    "CP_Gen_Name",
    "UNIFORMAT_CODE",
    "ALL_MODEL_TYPE_COMMENTS",
    "ALL_MODEL_DESCRIPTION",
    "ALL_MODEL_MANUFACTURER",
    "ALL_MODEL_MODEL"
]

Categories_Dict = {
    'FF': [BuiltInCategory.OST_Furniture, '02_FURNITURE'],
    'CF': [BuiltInCategory.OST_Furniture, '02_FURNITURE'],
    'MF': [BuiltInCategory.OST_GenericModel, '02_FURNITURE'],
    'PLF': [BuiltInCategory.OST_PlumbingFixtures, '06_PLUMBING_FIXTURES'],
    'SE': [BuiltInCategory.OST_SpecialityEquipment, '03_SPECIALTY_EQUIPMENT'],
    'EF': [BuiltInCategory.OST_ElectricalFixtures, '05_ELECTRICAL_FIXTURES'],
    'LF01': [BuiltInCategory.OST_LightingFixtures, '04_LIGHTING_FIXTURES']
}


#Собираем все параметры семейства в лист


fam_type_input_names = [
    "CP_Gen_Mark",
    "ALL_MODEL_TYPE_COMMENTS",
    "CP_Dim_Length",
    "CP_Dim_Width",
    "CP_Dim_Height"
]

marks_modified = []
for mark in selected_marks:
    #Определяем общую марку семейства
    dot_ind = mark.find('.')
    slash_ind = mark.find('-')
    # print slash_ind
    if dot_ind != -1 and slash_ind == -1:
        mark_mod = mark[:dot_ind]
    else:
        mark_mod = mark
    marks_modified.append(mark_mod)

lib_ind = excel_path.find('01_БИБЛИОТЕКА_ПРОЕКТА')
library_folder_path = excel_path[:lib_ind+len('01_БИБЛИОТЕКА_ПРОЕКТА')+1] + '01_СЕМЕЙСТВА\\'

family_names = []
family_dict = {}
for target_value in marks_modified:
    # print '==================ЧТЕНИЕ ДАННЫХ ИЗ РЕЕСТРА========================='
    
    #Извлекаем значения листа и строк типоразмеров, соотвествующих марке семейства
    mark_row_list = []
    mark_row_list = find_cell_coordinates(excel_path, target_value)[0]
    multiple_types = False
    if isinstance(mark_row_list, list):
        multiple_types = True
    if multiple_types:
        sheet_index = find_cell_coordinates(excel_path, target_value)[2][0]
    else:
        sheet_index = find_cell_coordinates(excel_path, target_value)[2]
    target_sheet = workbook.sheet_by_index(sheet_index)
    # exec("print 'Код каталога в реестре - '+ '{}'".format(target_sheet.name))

    #Собираем все значения параметров из реестра в список
    types_dict = {}
    param_dict = {}
    if multiple_types:
        for row_ind in mark_row_list:
            param_dict = {}
            for param in Parameter_List_Internal:
                sheet_ind_list = find_cell_coordinates(excel_path, param)[2]
                if isinstance(sheet_ind_list, list):
                    for x, ind in enumerate(sheet_ind_list):
                        if x == sheet_index:
                            col_ind = find_cell_coordinates(excel_path, param)[1][ind]-1
                            # exec("print 'Координаты параметра типа - ' + '{}' + ':' + '{}'".format(row_ind,col_ind))
                            param_dict[param] = target_sheet.cell_value(row_ind-1, col_ind)
            sheet_ind_list = find_cell_coordinates(excel_path, "CP_Gen_Mark")[2]
            if isinstance(sheet_ind_list, list):
                for x, ind in enumerate(sheet_ind_list):
                    if x == sheet_index:
                        col_ind = find_cell_coordinates(excel_path, "CP_Gen_Mark")[1][ind]-1
                        param_dict["CP_Gen_Mark"] = target_sheet.cell_value(row_ind-1, col_ind)
            sheet_ind_list = find_cell_coordinates(excel_path, "Discipline")[2]
            if isinstance(sheet_ind_list, list):
                for x, ind in enumerate(sheet_ind_list):
                    if x == sheet_index:
                        col_ind = find_cell_coordinates(excel_path, "Discipline")[1][ind]-1
                        param_dict["Discipline"] = target_sheet.cell_value(row_ind-1, col_ind)
                        check_list = [param_dict[i] for i in param_dict.keys()]
    else:
        row_ind = mark_row_list
        # print "check3"
        for param in Parameter_List_Internal:
            sheet_ind_list_2 = find_cell_coordinates(excel_path, param)[2]
            # print sheet_ind_list_2
            if isinstance(sheet_ind_list_2, list):
                # print "pass"
                for x, ind in enumerate(sheet_ind_list_2):
                    # print "pass2"
                    # print x
                    # print sheet_index
                    if x == sheet_index-1:
                        # print "pass3"
                        # print ind
                        col_ind = find_cell_coordinates(excel_path, param)[1][ind-3]-1
                        # print col_ind
                        param_dict[param] = target_sheet.cell_value(row_ind-1, col_ind)
                # print param_dict['CP_Dim_Width']
        sheet_ind_list_2 = find_cell_coordinates(excel_path, "CP_Gen_Mark")[2]
        if isinstance(sheet_ind_list_2, list):
            for x, ind in enumerate(sheet_ind_list_2):
                if x == sheet_index-1:
                    col_ind = find_cell_coordinates(excel_path, "CP_Gen_Mark")[1][ind-3]-1
                    param_dict["CP_Gen_Mark"] = target_sheet.cell_value(row_ind-1, col_ind)
        sheet_ind_list_2 = find_cell_coordinates(excel_path, "Discipline")[2]
        if isinstance(sheet_ind_list_2, list):
            for x, ind in enumerate(sheet_ind_list_2):
                if x == sheet_index-1:
                    col_ind = find_cell_coordinates(excel_path, "Discipline")[1][ind-3]-1
                    param_dict["Discipline"] = target_sheet.cell_value(row_ind-1, col_ind)
                    check_list = [param_dict[i] for i in param_dict.keys()]
 

    fam_category = Categories_Dict[target_sheet.name][0]

    #Составляем имя файла семейства по-умолчанию
    prefix_ind_start = find_all_indexes(excel_path,'\\')[-1]+1
    excel_file_name = excel_path[prefix_ind_start:]
    prefix_ind_end = excel_file_name.find('_')
    prefix = excel_file_name[:prefix_ind_end]
    Gen_Mark_ind = param_dict["CP_Gen_Mark"].find('.')
    if Gen_Mark_ind == -1:
        Gen_Mark_modified = param_dict["CP_Gen_Mark"]
    else:
        Gen_Mark_modified = param_dict["CP_Gen_Mark"][:Gen_Mark_ind]
    revit_version = app.VersionNumber
    family_name = prefix + '_' + param_dict["Discipline"] + '_' + param_dict["UNIFORMAT_CODE"] + '_' + Gen_Mark_modified + '_' + param_dict["ALL_MODEL_TYPE_COMMENTS"] + '_R' + revit_version[-2:] + '.rfa'
    index = excel_path.find('01_БИБЛИОТЕКА_ПРОЕКТА')
    family_file_path = excel_path[:index+len('01_БИБЛИОТЕКА_ПРОЕКТА')+1] + '01_СЕМЕЙСТВА\\' + Categories_Dict[target_sheet.name][1] + '\\'
    # print family_file_path



    #Записываем значения параметров в семействах
    # print library_folder_path
    # print family_name
    fam_path = False
    fam_exist_in_lib = False
    for root, dirs, files in os.walk(library_folder_path):
        # print 'check1'
        for file in files:
            # print file
            # print 'check2'
            if file.endswith(".rfa") and '.00' not in file and family_name in file:
                # print 'check3'
                fam_path = os.path.join(root, file)
                fam_exist_in_lib = True
                # print fam_path
    # print fam_path
    if not fam_path:
        fam_path = os.path.join(family_file_path, family_name)
        # print fam_path
    if family_name not in family_dict.keys():
        family_dict[family_name] = []
    family_dict[family_name].append(fam_path)
    family_dict[family_name].append(fam_exist_in_lib)

print '================================================================================'
print '================================================================================'
old_doc = False
with forms.ProgressBar(step = 1, title = 'Загрузка семейств из библиотеки..' + '{value} из {max_value}', cancellable = True) as pb:

    pbTotal1 = len(family_dict.keys())
    pbCount1 = 1
    pbWorks1 = 0

    loadOption = FamilyOption()

    with revit.Transaction('Загрузка семейств'):
        for fam_name, fam_data in family_dict.items():
            doc_path = fam_data[0].replace("\\", "\\\\")
            fam_exist_in_lib_= fam_data[1]
            if fam_exist_in_lib_:
                exec("print 'Семейство обнаружено - '+ '{}'".format(fam_name))
                exec("print 'Путь к семейству: '+ '{}'".format(str(doc_path)))
                if pb.cancelled:
                    break
                else:
                    pbWorks1 += family_reload(doc_path,loadOption, doc)
                pb.update_progress(pbCount1, pbTotal1)
                pbCount1 += 1
            else:
                exec("print 'Семейство не обнаружено - '+ '{}'".format(fam_name))
                exec("print 'Путь к семейству: '+ '{}'".format(str(doc_path)))
msg = str(pbWorks1) + '/' + str(pbTotal1) + ' семейств обновлены'
forms.alert(msg, title = 'Загрузка завершена', warn_icon = False)