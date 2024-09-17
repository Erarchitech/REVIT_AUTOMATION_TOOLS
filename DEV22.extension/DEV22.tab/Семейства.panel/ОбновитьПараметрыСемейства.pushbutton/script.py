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
from Autodesk.Revit.UI import TaskDialog
from RevitServices.Persistence import DocumentManager
clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)
clr.AddReference('RevitAPI')
from RevitServices.Transactions import TransactionManager
from System.IO import File

import pyrevit
from pyrevit import forms
from pyrevit import HOST_APP

uiapp = __revit__ # noqa
app = __revit__.Application # noqa
uidoc = __revit__.ActiveUIDocument # noqa
f_doc = __revit__.ActiveUIDocument.Document # noqa


import clr
import sys
# sys.path.append(r'X:\00_РЕСУРСЫ\01_ЦЕНТРАЛЬНАЯ_БИБЛИОТЕКА\04_СКРИПТЫ\02_DEV\02_PYTHON')
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
clr.AddReference('RevitNodes')
clr.AddReference("System.Windows.Forms")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import FilteredElementCollector as FEC
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

#USER INPUTS
excel_path = pyrevit.forms.pick_file(file_ext='xls', files_filter='', init_dir='', restore_dir=True, multi_file=False, unc_paths=False, title='Выбери файл Excel Реестра')
workbook = xlrd.open_workbook(excel_path)


# sheets = [workbook.sheet_by_index(sheet_index) for sheet_index in range(workbook.nsheets)]
# sheets = sheets[1:]

# marks = []
# type_row_dict = {}
# # for sheet in selected_sheets:
# #     # Iterate through all rows and columns
# #     for row in range(sheet.nrows):
# #         for col in range(sheet.ncols):
# #             cell_value = str(sheet.cell_value(row, col))
# #             # Check if the cell contains the target value
# #             if cell_value == 'CP_Gen_Mark':
# #                 for row in range(sheet.nrows):
# #                     mark = str(sheet.cell_value(row, col))
# #                     marks.append(mark)
# elements_to_exclude = ["CP_Gen_Mark", 'Марка типа', 'Марка', '']
# marks = [element for element in marks if element not in elements_to_exclude]
# selected_marks = pyrevit.forms.SelectFromList.show(marks, title = 'Выбери Марки', multiselect = True, button_name = 'Выбрать')

# if len(selected_marks) == 0:
#     selected_marks = marks

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
    "ALL_MODEL_MODEL",
    'CP_Gen_Link',
    'CP_Gen_Image'
]

Categories_Dict = {
    'FF': [BuiltInCategory.OST_Furniture, '02_FURNITURE'],
    'CF': [BuiltInCategory.OST_Furniture, '02_FURNITURE'],
    'MF': [BuiltInCategory.OST_GenericModel, '02_FURNITURE'],
    'PLF': [BuiltInCategory.OST_PlumbingFixtures, '06_PLUMBING_FIXTURES'],
    'SE': [BuiltInCategory.OST_SpecialityEquipment, '03_SPECIALTY_EQUIPMENT'],
    'EF': [BuiltInCategory.OST_ElectricalFixtures, '05_ELECTRICAL_FIXTURES']
}


#Собираем все параметры семейства в лист


fam_type_input_names = [
    "CP_Gen_Mark",
    "ALL_MODEL_TYPE_COMMENTS",
    "CP_Dim_Length",
    "CP_Dim_Width",
    "CP_Dim_Height"
]


# doc_path = doc.PathName
#Считываем значение параметра CP_Gen_Mark для поиска других параметров в реестре
f_manager = f_doc.FamilyManager
fam_parameter = f_manager.Parameters.ForwardIterator()
fam_parameter_list = []
fam_param_names = []
fam_types = f_manager.Types
# fam_types_iterator = fam_types.ForwardIterator()
fam_type_list = []
# while fam_types_iterator.MoveNext():
#     fam_type_list.append(fam_types_iterator.Current)
# type_is_single = False
# if fam_types.Size == 1:
#     type_is_single = True
# current_type = f_manager.CurrentType

while fam_parameter.MoveNext():
    fam_param = fam_parameter.Current
    fam_parameter_list.append(fam_param)
    fam_param_names.append(fam_param.Definition.BuiltInParameter)
    par_name = fam_param.Definition.Name
    if par_name == 'CP_Gen_Mark':
        Gen_Mark_param = fam_param
    if par_name == 'CP_Gen_Ver':
        Gen_Ver_param = fam_param

current_type = f_manager.CurrentType
gen_mark_value = current_type.AsString(Gen_Mark_param)

marks_modified = []

#Определяем общую марку семейства
dot_ind = gen_mark_value.find('.')
slash_ind = gen_mark_value.find('-')
# print slash_ind
if dot_ind != -1 and slash_ind == -1:
    family_mark_mod = gen_mark_value[:dot_ind]
else:
    family_mark_mod = gen_mark_value

lib_ind = excel_path.find('01_БИБЛИОТЕКА_ПРОЕКТА')
library_folder_path = excel_path[:lib_ind+len('01_БИБЛИОТЕКА_ПРОЕКТА')+1] + '01_СЕМЕЙСТВА\\'


target_value = family_mark_mod
print '==================ЧТЕНИЕ ДАННЫХ ИЗ РЕЕСТРА========================='
print str('Марка семейства - ' + target_value)

#Извлекаем значения листа и строк типоразмеров, соотвествующих марке семейства
mark_row_list = []
mark_row_list = find_cell_coordinates(excel_path, target_value)[0]
multiple_types = False
# print 'check03'
# print excel_path
if isinstance(mark_row_list, list) and len(mark_row_list)>1:
    multiple_types = True
    # print multiple_types
if multiple_types:
    # print 'check04'
    sheet_index = find_cell_coordinates(excel_path, target_value)[2][0]
elif isinstance(mark_row_list, list) and len(mark_row_list) == 1:
    # print 'check06'
    sheet_index = find_cell_coordinates(excel_path, target_value)[2][0]
else:
    # print 'check05'
    sheet_index = find_cell_coordinates(excel_path, target_value)[2]
# print sheet_index
target_sheet = workbook.sheet_by_index(sheet_index)
sheet = target_sheet
# print len(sheet)
exec("print 'Код каталога в реестре - '+ '{}'".format(target_sheet.name))

#Собираем все значения параметров из реестра в список
types_dict = {}
param_dict = {}
# print 'check01'
if multiple_types:
    for row_ind in mark_row_list:
        print '_________________________________________'
        param_dict = {}
        for param in Parameter_List_Internal:
            # print param
            # print 'check02'
            sheet_ind_list = find_cell_coordinates(excel_path, param)[2]
            # print sheet_ind_list
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

        # param_dict = {}
        # print '    '
        # Составляем имя типоразмера из значений параметров
        fam_type_input_values = []
        for name in fam_type_input_names:
            for p,v in param_dict.items():
                # print p
                if p == name:
                    if isinstance(v, float):
                        if v > 0:
                            fam_type_input_values.append(int(v))
                        else:
                            fam_type_input_values.append('')
                    else:
                        fam_type_input_values.append(v)
        # print fam_type_input_values
        type_name = fam_type_input_values[0] + "_" \
            + fam_type_input_values[1] + "_" \
            + str(fam_type_input_values[2]) + "x" \
            + str(fam_type_input_values[3]) + "x" \
            + str(fam_type_input_values[4]) + "(h)"
        # print type_name
        
        types_dict[type_name] = param_dict
        check_list = [param_dict[i] for i in param_dict.keys()]

        exec("print 'Обнаружен тип - ' + '{}'".format(type_name))
        exec("print 'Строка в реестре - '+ '{}'".format(row_ind))
        exec("print 'Количество считанных параметров: ' + '{}'".format(len(check_list)))
    #     print check_list
    # print "check2"
    # type_name_list = [name for name in types_dict.keys()]
else:
    print '_________________________________________'
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
                # print check_list
    
    # print 'check0'
    # param_dict = {}
    # print '    '
    # Составляем имя типоразмера из значений параметров
    fam_type_input_values = []
    for name in fam_type_input_names:
        for p,v in param_dict.items():
            # print p
            if p == name:
                # print name
                if isinstance(v, float):
                    if v > 0:
                        fam_type_input_values.append(int(v))
                    else:
                        fam_type_input_values.append('')
                else:
                    fam_type_input_values.append(v)
    # print fam_type_input_values
    # print 'check01'
    # print fam_type_input_values
    type_name = fam_type_input_values[0] + "_" \
        + fam_type_input_values[1] + "_" \
        + str(fam_type_input_values[2]) + "x" \
        + str(fam_type_input_values[3]) + "x" \
        + str(fam_type_input_values[4]) + "(h)"
    # print type_name
    
    types_dict[type_name] = param_dict
    # check_list = [param_dict[i] for i in param_dict.keys()]
    # print check_list
    print str('Обнаружен тип - ' + type_name)
    exec("print 'Категория - '+ '{}'".format(target_sheet.name))
    exec("print 'Строка в реестре - '+ '{}'".format(row_ind))
    exec("print 'Количество считанных параметров: ' + '{}'".format(len(check_list)))
# print "checkout"

#Проверка прохода по типам и параметрам
for type_, par_dict in types_dict.items():
    print type_
    print "Type Parameters:"
    for par, value in par_dict.items():
        exec("print '{} - ' + '{}'".format(par, value))
        print('')


fam_category = Categories_Dict[target_sheet.name][0]

#Составляем имя файла семейства по-умолчанию
prefix_ind_start = find_all_indexes(excel_path,'\\')[-1]+1
excel_file_name = excel_path[prefix_ind_start:]
prefix_ind_end = excel_file_name.find('_')
prefix = excel_file_name[:prefix_ind_end]
Gen_Mark_ind = param_dict["CP_Gen_Mark"].find('.')
if Gen_Mark_ind == -1:
    Gen_family_Mark_modified = param_dict["CP_Gen_Mark"]
else:
    Gen_family_Mark_modified = param_dict["CP_Gen_Mark"][:Gen_Mark_ind]
revit_version = app.VersionNumber
family_name = prefix + '_' + param_dict["Discipline"] + '_' + param_dict["UNIFORMAT_CODE"] + '_' + Gen_family_Mark_modified + '_' + param_dict["ALL_MODEL_TYPE_COMMENTS"] + '_R' + revit_version[-2:] + '.rfa'
index = excel_path.find('01_БИБЛИОТЕКА_ПРОЕКТА')
family_file_path = excel_path[:index+len('01_БИБЛИОТЕКА_ПРОЕКТА')+1] + '01_СЕМЕЙСТВА\\' + Categories_Dict[target_sheet.name][1] + '\\'
fam_path = os.path.join(family_file_path, family_name)
# print family_file_path

# for fam_n in family_dict.keys():
#     print fam_n
# print 'check'

# doc_paths = iterate_files(library_folder_path, family_names)
# print doc_paths
# for doc_path in doc_paths:
#     print doc_path
old_doc = False

print '================================================================================'
print '================================================================================'

# print fam_name
# print library_folder_path
print '=========================ЗАПИСЬ ДАННЫХ В СЕМЕЙСТВО========================='

update_types_dict = types_dict.copy()

# for t in types_dict.keys():
#     print t

print '__________ТРАНЗАКЦИЯ ОТКРЫТА________'
with Transaction(f_doc, 'Обновление семейства') as t:
    t.Start()
    for fam_type in fam_types:
        f_manager.CurrentType = fam_type
        gen_mark_value = fam_type.AsString(Gen_Mark_param)
        type_found = False
        if fam_types.Size > 0:
            for type_, par_dict in update_types_dict.items():
                values = []
                for value in par_dict.values():
                        values.append(value)
                # print values
                if gen_mark_value in values:
                    type_found = True
                    # print 'check_last'
                    try:
                        # print 'check_99'
                        # current_type.Name = type_
                        print('')
                        exec("print 'Обновление имени типа с - {} '+ '- {}'".format(f_manager.CurrentType.Name,type_))
                        # current_type.Name = type_
                        new_type_name = f_manager.RenameCurrentType(str(type_))
                        print str("Имя текущего типа обновлено - " + type_)
                    except:
                        print 'Имя текущего типоразмера соответствует реестру'
                    for ind, par_new in enumerate(Parameter_List_Internal):
                        for par_exist in fam_parameter_list:
                            if par_exist.Definition.BuiltInParameter != BuiltInParameter.INVALID:
                                par_exist_name = BuiltInParameter.ToString(par_exist.Definition.BuiltInParameter)
                            else:
                                par_exist_name = par_exist.Definition.Name
                            if par_new == 'CP_Gen_Image' and par_new == par_exist_name:
                                # print par_exist.StorageType
                                # Use ImageTypeOptions to load the image file into Revit
                                
                                image_path ="X:\\" + par_dict[par_new]
                                if not File.Exists(image_path):
                                    print("Файл изображения не найден!")
                                else:
                                    # print image_path
                                    try:
                                        # Ensure valid ImageTypeOptions setup and proper image format
                                        image_options = ImageTypeOptions(image_path, False, ImageTypeSource.Import)
                                        image_type = ImageType.Create(f_doc, image_options)
                                        # print image_type
                                        if image_type is None:
                                            print("Ошибка, не удалось загрузить изображение.")
                                            t.RollBack()
                                        else:
                                            f_manager.Set(
                                            par_exist,
                                            image_type.Id
                                            )
                                            print("Изображение добавлено успешно.")
                                    except Exception as e:
                                        print("Error", str(e))
                                        t.RollBack()
                                    
                            elif par_new == par_exist_name:
                                storage_type = par_exist.StorageType
                                if storage_type == StorageType.Double:
                                    f_manager.Set(
                                    par_exist,
                                    UnitUtils.ConvertToInternalUnits(par_dict[par_new],UnitTypeId.Millimeters)
                                    )
                                else:
                                    f_manager.Set(
                                    par_exist,
                                    par_dict[par_new]
                                    )
                                # print str(par_new + ' - ' + str(par_dict[par_new]))
                            
                    del update_types_dict[type_]
            if type_found == False:
                exec("print 'Удаление типа - '+ '{}'".format(f_manager.CurrentType.Name))
                f_manager.DeleteCurrentType()

        for type_, par_dict in update_types_dict.items():
            try:
                f_manager.NewType(type_)
                print str("Добавлен новый тип - " + type_)
                for ind, par_new in enumerate(Parameter_List_Internal):
                    for par_exist in fam_parameter_list:
                        if par_exist.Definition.BuiltInParameter != BuiltInParameter.INVALID:
                            par_exist_name = BuiltInParameter.ToString(par_exist.Definition.BuiltInParameter)
                        else:
                            par_exist_name = par_exist.Definition.Name
                        if par_new == par_exist_name:
                            storage_type = par_exist.StorageType
                            if storage_type == StorageType.Double:
                                f_manager.Set(
                                par_exist,
                                UnitUtils.ConvertToInternalUnits(par_dict[par_new],UnitTypeId.Millimeters)
                                )
                            else:
                                f_manager.Set(
                                par_exist,
                                par_dict[par_new]
                                )
                            # print str(par_new + ' - ' + str(par_dict[par_new]))
            except:
                pass

        print 'Параметры записаны'
    f_doc.OwnerFamily.FamilyCategory = Category.GetCategory(f_doc, fam_category)
    print str('Категория семейства обновлена - ' + Category.GetCategory(f_doc, fam_category).Name)
    f_doc.OwnerFamily.Parameter[BuiltInParameter.ROOM_CALCULATION_POINT].Set(1)
    print str('Точка вставки обновлена')
    f_doc.OwnerFamily.Parameter[BuiltInParameter.FAMILY_SHARED].Set(1)
    print str('Семейство является общим')
    current_date = datetime.now()
    formatted_date = current_date.strftime("%y.%m.%d")
    f_manager.SetFormula(Gen_Ver_param, str('"' + 'V'+formatted_date + '"'))
    print str('Версия семейства обновлена от ' + formatted_date)
    
    # Шляпа с удалением типов с неопределенными марками не работает пока
    f_doc.Regenerate()
    values_after = []
    for fam_type_after in f_manager.Types:
        f_manager.CurrentType = fam_type_after
        fam_type_match = False
        # print fam_type_after.Name
        # print "||||"
        # print fam_type_after.AsString(Gen_Mark_param)
        for type_after in types_dict.keys():
            print type_after
            if fam_type_after.Name == type_after:
                # print 'Тип соответствует'
                fam_type_match = True
        if not fam_type_match:
            exec("print 'Удаление типа - '+ '{}'".format(f_manager.CurrentType.Name))
            f_manager.DeleteCurrentType()
        # print '----------------------'
    # Close the family document
    # close_inactive_docs(uiapp)
    # TransactionManager.Instance.ForceCloseTransaction()
    f_doc.OwnerFamily.FamilyCategory = Category.GetCategory(f_doc, fam_category)
    print str('Категория семейства обновлена - ' + Category.GetCategory(f_doc, fam_category).Name)
    f_doc.OwnerFamily.Parameter[BuiltInParameter.ROOM_CALCULATION_POINT].Set(1)
    print str('Точка вставки обновлена')
    f_doc.OwnerFamily.Parameter[BuiltInParameter.FAMILY_SHARED].Set(1)
    print str('Семейство является общим')
    current_date = datetime.now()
    formatted_date = current_date.strftime("%y.%m.%d")
    f_manager.SetFormula(Gen_Ver_param, str('"' + 'V'+formatted_date + '"'))
    print str('Версия семейства обновлена от ' + formatted_date)
    t.Commit()
        
    # Save the family document
    save_as_options = SaveAsOptions()
    save_as_options.OverwriteExistingFile = True
    view = FEC(f_doc).OfClass(View3D).FirstElement()
    save_as_options.PreviewViewId = view.Id
    f_doc.SaveAs(fam_path, save_as_options)
    exec("print 'Семейство сохранено: '+ '{}'".format(f_doc.Title))

ind = excel_path.find('01_БИБЛИОТЕКА_ПРОЕКТА')
modified_path_str = excel_path.replace("\\",'\\\\')
family_file_path = excel_path[:ind+len('01_БИБЛИОТЕКА_ПРОЕКТА')+1] + '01_СЕМЕЙСТВА\\' + Categories_Dict[sheet.name][1] + '\\'

# print modified_path_str
# print family_file_path
# save_file = pyrevit.forms.save_file(file_ext='rfa', files_filter='', init_dir = family_file_path, default_name= family_name, restore_dir=True, unc_paths=False, title= 'Сохранить семейство в библиотеку проекта')