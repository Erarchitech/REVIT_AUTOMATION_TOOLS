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
from System.IO import File

import pyrevit
from pyrevit import forms
from pyrevit import HOST_APP
from pyrevit import revit

uiapp = __revit__ # noqa
app = __revit__.Application # noqa
uiapp = HOST_APP.uiapp
# uidoc = __revit__.ActiveUIDocument # noqa
doc = __revit__.ActiveUIDocument.Document # noqa

# Функция для получения типа по-умолчанию по имени класса
def get_default_type_by_class(doc, class_name):
    valid_class_list = [
        FloorType,
        WallType,
        CeilingType
    ]
    if class_name in valid_class_list:
        collector = FEC(doc).OfClass(class_name).WhereElementIsElementType().ToElements()
        # print('check_01')
        for type_ in collector:
            # print(type_)
            # print(type_.Parameter[BuiltInParameter.ALL_MODEL_TYPE_NAME].AsString())
            mark_param = type_.LookupParameter('CP_Gen_Mark')
            if mark_param:
                storage_type = mark_param.StorageType
                if storage_type == StorageType.String:
                    mark_value = mark_param.AsString()
                    if mark_value is not None:
                        if 'XX' in mark_value:
                            return type_
            else:
                print('параметр CP_Gen_Mark не считан')
    else:
        return None
    
# Функция для получения текущего типа в проекте по имени типа
def get_type_by_name(doc, class_name, type_name):
    collector = FEC(doc).OfClass(class_name)
    for el in collector:
        if el.Parameter[BuiltInParameter.ALL_MODEL_TYPE_NAME].AsString() == type_name:
            return el
    return None

# Функция для поиска материала по имени
def get_material_by_name(doc, material_name):
    materials = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Materials).WhereElementIsNotElementType().ToElements()
    for material in materials:
        if material.Name == material_name:
            return material.Id
    return None

# Функция для создания нового материала
def create_new_material(doc, material_name):
    new_material_id = Material.Create(doc, material_name)
    return new_material_id

# Функция для получения значения параметра по имени
def get_parameter_value_by_name(element, param_name):
    param = element.LookupParameter(param_name)  # Поиск параметра по имени
    if param:
        # Проверка типа параметра и получение значения
        if param.StorageType == Autodesk.Revit.DB.StorageType.String:
            return param.AsString()
        elif param.StorageType == Autodesk.Revit.DB.StorageType.Double:
            return param.AsDouble()  # Значение возвращается в футах для длин
        elif param.StorageType == Autodesk.Revit.DB.StorageType.Integer:
            return param.AsInteger()
        elif param.StorageType == Autodesk.Revit.DB.StorageType.ElementId:
            return param.AsElementId()
    return None

# Функция для извлечения индекса из ключа
def extract_index(key):
    # Разделяем строку по символу "_" и берем последнюю часть как индекс
    return int(key.split('_')[-1])

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
    'CP_Gen_Mark',
    'CP_Gen_Image'
]

# ВВОД списков материалов и их толщин для добавления в типоразмер

def parameters_updater(doc, type_dict):
    updated_types = []
    # Начало транзакции
    with Transaction(doc, 'Обновление параметров типов системных семейств') as t:
        t.Start()

        for type_name, param_dict in type_dict.items():
            fam_category = param_dict['Category']
            # Находим сущетсвующий типоразмер семейства соответствующий имени типа из словаря
            collector = FEC(doc).OfClass(fam_category)
            for exist_type in collector:
                if exist_type.Parameter[BuiltInParameter.ALL_MODEL_TYPE_NAME].AsString() == type_name:
                    update_type = get_type_by_name(doc, fam_category, type_name)
                    updated_types.append(update_type.Parameter[BuiltInParameter.ALL_MODEL_TYPE_NAME].AsString())

                    for par_new in Parameter_List_Internal:
                        # Перебираем все параметры данного типа
                        for par_exist in update_type.Parameters:
                            # param_name = par_exist.Definition.Name
                            # param_value = par_exist.AsValueString() or par_exist.AsString()
                            # print("Параметр: {}, Значение: {}".format(param_name, param_value))
                            # print('check01')
                            if par_exist.Definition.BuiltInParameter != BuiltInParameter.INVALID:
                                par_exist_name = BuiltInParameter.ToString(par_exist.Definition.BuiltInParameter)
                            else:
                                par_exist_name = par_exist.Definition.Name
                            if par_new == 'CP_Gen_Image' and par_new == par_exist_name:
                                # print par_exist.StorageType
                                # Use ImageTypeOptions to load the image file into Revit
                                
                                image_path ="X:\\" + param_dict[par_new]
                                if not File.Exists(image_path):
                                    print("Файл изображения не найден!")
                                else:
                                    # print image_path
                                    try:
                                        # Ensure valid ImageTypeOptions setup and proper image format
                                        image_options = ImageTypeOptions(image_path, False, ImageTypeSource.Import)
                                        image_type = ImageType.Create(doc, image_options)
                                        # print image_type
                                        if image_type is None:
                                            print("Ошибка, не удалось загрузить изображение.")
                                            t.RollBack()
                                        else:
                                            par_exist.Set(image_type.Id)
                                            print("Изображение добавлено успешно.")
                                    except Exception as e:
                                        print("Error", str(e))
                                        t.RollBack()
                                    
                            elif par_new == 'CP_Gen_Mark' and par_new == par_exist_name:
                                par_exist.Set(param_dict[par_new])
                                print('{} обновлен успешно: {}'.format(par_new, param_dict[par_new]))
                            elif par_new == par_exist_name:
                                try:
                                    par_value = param_dict[par_new][1]
                                except:
                                    par_value = None
                                if par_value:
                                    par_exist.Set(param_dict[par_new][1])
                                    print('{} обновлен успешно: {}'.format(par_new, param_dict[par_new][1]))
                        print('____________________')
            
        t.Commit()
    return  updated_types
    # Создание новой составной структуры
    # typed_layers: typing.List[CompoundStructureLayer] = layers
    # new_compound_structure = CompoundStructure.CreateSingleLayerCompoundStructure(MaterialFunctionAssignment.Structure, sorted_param_dict['param01'][1], material_id)

    # Применение новой структуры к новому типу перекрытия
    # print(layers)
    
    # print('check02')
    
    # print('check03')

        

    # Завершение транзакции
    # t.Commit()