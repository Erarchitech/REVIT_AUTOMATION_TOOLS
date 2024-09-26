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

def types_updater(doc, type_dict):
    #ПРЕДВАРИТЕЛЬНАЯ ОБРАБОТКА ВВОДНЫХ ДАННЫХ
    #Конвертация единиц толщин в системные
    updated_types = []
    for param_dict in type_dict.values():
        for key, v in param_dict.items():
            if 'CP_Mat_Finish_' in key:
                material_width = v[0]
                material_width_internal = UnitUtils.ConvertToInternalUnits(material_width,UnitTypeId.Millimeters)
                v[0] = material_width_internal

    # Начало транзакции
    with Transaction(doc, 'Обновление типов системных семейств из реестра') as t:
        t.Start()

        for new_type_name, param_dict in type_dict.items():
            # print(param_dict['Category'])
            # Получение существующего типа перекрытия (например, "Generic 12")
            default_type = get_default_type_by_class(doc, param_dict['Category'])
            # Получение существующей составной структуры перекрытия
            compound_structure = default_type.GetCompoundStructure()
            # Получение индексов слоев сердцевины (Core) для модификации структуры
            first_core_layer = compound_structure.GetFirstCoreLayerIndex()  # Индекс первого слоя сердцевины
            last_core_layer = compound_structure.GetLastCoreLayerIndex()    # Индекс последнего слоя сердцевины

            #Поиск текущего типоразмера в проекте
            found_type = get_type_by_name(doc, param_dict['Category'], new_type_name)
            
            key_list = []
            for key in param_dict.keys():
                if 'CP_Mat_Finish_' in key:
                    key_list.append(key)
            sorted_key_list = sorted(key_list)
            # print(sorted_key_list)
            
            if found_type is None:
                # Создание нового типа 
                new_type = default_type.Duplicate(new_type_name)
                print new_type
                print('Создан новый типоразмер - {}'.format(new_type_name))
                layers = []

                # Поиск материала по имени
                for k in sorted_key_list:
                    material_name = param_dict[k][2]
                    material_width = param_dict[k][0]
                    print(material_name)
                    material_id = get_material_by_name(doc, material_name)
                    # Если материал не найден, создаем новый
                    if material_id is None:
                        material_id = create_new_material(doc, material_name)
                        print('Создан новый материал - {}'.format(material_name))
                    else:
                        # print('Данный материал уже существует в проекте - {}'.format(material_name))
                        pass

                    # Создание составной структуры
                    # Здесь задается толщина каждого слоя и соответствующий материал (ElementId)
                    new_layer = CompoundStructureLayer(material_width, MaterialFunctionAssignment.Structure, material_id)
                    # Добавление нового слоя в список слоев
                    layers.append(new_layer)
                    print('Слой добавлен')
                # print('check02')
                new_structure = CompoundStructure.CreateSimpleCompoundStructure(layers)
                new_structure.EndCap = EndCapCondition.NoEndCap
                new_type.SetCompoundStructure(new_structure)
                print new_type
                print('Новый типоразмер добавлен в проект - {}'.format(new_type_name))
                print('__')
                updated_types.append(new_type)
            else:
                # Обновление существующего типа
                print('Найден существующий типоразмер - {}'.format(found_type.Parameter[BuiltInParameter.ALL_MODEL_TYPE_NAME].AsString()))
                cs = found_type.GetCompoundStructure()
                exist_layers = cs.GetLayers()
                check = 0
                for ind, k in enumerate(sorted_key_list):

                    new_material_name = param_dict[k][2]
                    new_material_width = param_dict[k][0]
                    new_material_id = get_material_by_name(doc, new_material_name)

                    # Если материал не найден, создаем новый
                    if new_material_id is None:
                        new_material_id = create_new_material(doc, new_material_name)
                        print('Создан новый материал - {}'.format(new_material_name))
                    else:
                        # print('Данный материал уже существует в проекте - {}'.format(new_material_name))
                        pass

                    if ind < len(exist_layers):

                        exist_material_id = exist_layers[ind].MaterialId
                        exist_material_width = exist_layers[ind].Width

                        if new_material_id != exist_material_id:
                            print(new_material_name)
                            exist_layers[ind].MaterialId = new_material_id
                            print('Исправлен материал в пироге с индексом {}'.format(ind+1))
                            check+=1
                        if new_material_width != exist_material_width:
                            print(new_material_name)
                            print(new_material_width)
                            exist_layers[ind].Width = new_material_width
                            print('Перезаписана толщина слоя в пироге с индексом {}: {}'.format(ind+1, UnitUtils.ConvertFromInternalUnits(new_material_width,UnitTypeId.Millimeters)))
                            check+=1
                    else:
                        print(new_material_name)
                        new_layer = CompoundStructureLayer(new_material_width, MaterialFunctionAssignment.Structure, new_material_id)
                        exist_layers.Add(new_layer)
                        print('Добавлен новый слой в пироге - {} с индексом {}'.format(doc.GetElement(new_material_id).Name, ind+1))
                        check+=1
                if check == 0:
                    print('Пирог типа соотвествует реестру')
                else:
                    updated_cs = CompoundStructure.CreateSimpleCompoundStructure(exist_layers)
                    updated_cs.EndCap = EndCapCondition.NoEndCap
                    found_type.SetCompoundStructure(updated_cs)
                    print('Существующий типоразмер обновлен - {}'.format(found_type.Parameter[BuiltInParameter.ALL_MODEL_TYPE_NAME].AsString()))
                print('____________________')
                updated_types.append(found_type)
    # Создание новой составной структуры
    # typed_layers: typing.List[CompoundStructureLayer] = layers
    # new_compound_structure = CompoundStructure.CreateSingleLayerCompoundStructure(MaterialFunctionAssignment.Structure, sorted_param_dict['param01'][1], material_id)

    # Применение новой структуры к новому типу перекрытия
    # print(layers)
    
    # print('check02')
    
    # print('check03')

        

    # Завершение транзакции
        t.Commit()
    return  updated_types