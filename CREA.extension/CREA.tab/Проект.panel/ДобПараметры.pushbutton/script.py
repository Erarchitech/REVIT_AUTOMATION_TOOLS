# -*- coding: utf-8 -*-
import clr
import sys
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('RevitServices')
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import FilteredElementCollector as FEC

uiapp = __revit__ # noqa
app = __revit__.Application # noqa
uidoc = __revit__.ActiveUIDocument # noqa
doc = __revit__.ActiveUIDocument.Document # noqa



import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")
from Autodesk.Revit.DB import *
from RevitServices.Persistence import DocumentManager
from System.Collections.Generic import List

import pyrevit
from pyrevit import HOST_APP


uiapp = __revit__ # noqa
app = __revit__.Application # noqa
uiapp = HOST_APP.uiapp
# uidoc = __revit__.ActiveUIDocument # noqa
doc = __revit__.ActiveUIDocument.Document # noqa
app_version = __revit__.Application.VersionNumber # noqa

#Путь к ФОП, если он еще не подключен
if app.SharedParametersFilename == "":
    app.SharedParametersFilename = "\\Crea-Server\\Crea Plus\\00_РЕСУРСЫ\\01_ЦЕНТРАЛЬНАЯ_БИБЛИОТЕКА\\03_ШАБЛОНЫ\\CP_SHARED_PARAMETERS.txt"

# Открыть файл общих параметров
shared_params_file = app.OpenSharedParameterFile()
if shared_params_file is None:
    raise Exception("Не удалось открыть файл общих параметров.")

categories_set_A = [
                    BuiltInCategory.OST_Walls, 
                    BuiltInCategory.OST_Floors,
                    BuiltInCategory.OST_Roofs,
                    BuiltInCategory.OST_SpecialityEquipment,
                    BuiltInCategory.OST_ElectricalEquipment,
                    BuiltInCategory.OST_ElectricalFixtures,
                    BuiltInCategory.OST_Furniture,
                    BuiltInCategory.OST_LightingFixtures,
                    BuiltInCategory.OST_LightingDevices,
                    BuiltInCategory.OST_GenericModel,
                    BuiltInCategory.OST_Doors,
                    BuiltInCategory.OST_Windows,
                    BuiltInCategory.OST_Ceilings
                    ]

categories_set_B = [
                    BuiltInCategory.OST_SpecialityEquipment,
                    BuiltInCategory.OST_ElectricalEquipment,
                    BuiltInCategory.OST_ElectricalFixtures,
                    BuiltInCategory.OST_Furniture,
                    BuiltInCategory.OST_LightingFixtures,
                    BuiltInCategory.OST_LightingDevices,
                    BuiltInCategory.OST_GenericModel,
                    BuiltInCategory.OST_Doors,
                    BuiltInCategory.OST_Windows
                    ]

categories_set_С = [
                    BuiltInCategory.OST_Walls, 
                    BuiltInCategory.OST_Floors,
                    BuiltInCategory.OST_Roofs,
                    BuiltInCategory.OST_Ceilings
                    ]

categories_set_D = [
                    BuiltInCategory.OST_Sheets,
                    BuiltInCategory.OST_Views
                    ]


# Список параметров для добавления, представленный в виде словаря
parameters_to_add = {
    "CP_Gen_Mark": ["01_Gen_Общие", BuiltInParameterGroup.PG_TEXT, "Type", categories_set_A],
    "CP_Gen_Description": ["01_Gen_Общие", BuiltInParameterGroup.PG_TEXT, "Type", categories_set_A],
    "CP_Gen_Name": ["01_Gen_Общие", BuiltInParameterGroup.PG_TEXT,"Type", categories_set_A],
    "CP_Gen_Image": ["01_Gen_Общие", BuiltInParameterGroup.PG_GRAPHICS,"Type", categories_set_A],
    "CP_Dim_Height": ["03_Dim_Размеры", BuiltInParameterGroup.PG_GEOMETRY,"Type", categories_set_B],
    "CP_Dim_Length": ["03_Dim_Размеры", BuiltInParameterGroup.PG_GEOMETRY,"Type", categories_set_B],
    "CP_Dim_Width": ["03_Dim_Размеры", BuiltInParameterGroup.PG_GEOMETRY,"Type", categories_set_B],
    "CP_Gen_Ver": ["01_Gen_Общие", BuiltInParameterGroup.PG_IDENTITY_DATA,"Type", categories_set_A],
    "CP_Gen_SurfacePattern": ["01_Gen_Общие", BuiltInParameterGroup.PG_GRAPHICS,"Type", categories_set_С],
    "CP_Gen_Legend": ["01_Gen_Общие", BuiltInParameterGroup.PG_GRAPHICS,"Type", categories_set_С],
    "CP_Gen_Color": ["01_Gen_Общие", BuiltInParameterGroup.PG_TEXT,"Type", categories_set_A],
    "CP_Gen_Comments": ["01_Gen_Общие", BuiltInParameterGroup.PG_TEXT,"Type", categories_set_A],
    "CP_Gen_Construction_Scheme": ["01_Gen_Общие", BuiltInParameterGroup.PG_GRAPHICS,"Type", categories_set_С],
    "CP_Gen_RoomList": ["01_Gen_Общие", BuiltInParameterGroup.PG_TEXT,"Instance", categories_set_A],
    "CP_Gen_ShortName": ["01_Gen_Общие", BuiltInParameterGroup.PG_TEXT,"Type", categories_set_A],
    "CP_Gen_WallSide": ["01_Gen_Общие", BuiltInParameterGroup.PG_IDENTITY_DATA,"Instance", categories_set_B],
    "CP_Gen_Zone": ["01_Gen_Общие", BuiltInParameterGroup.PG_TEXT,"Instance", categories_set_A],
    "CP_Mat_Finish_01": ["02_Mat_Материалы", BuiltInParameterGroup.PG_MATERIALS,"Type", categories_set_С],
    "CP_Mat_Finish_02": ["02_Mat_Материалы", BuiltInParameterGroup.PG_MATERIALS,"Type", categories_set_С],
    "CP_Mat_Finish_03": ["02_Mat_Материалы", BuiltInParameterGroup.PG_MATERIALS,"Type", categories_set_С],
    "CP_Mat_Finish_04": ["02_Mat_Материалы", BuiltInParameterGroup.PG_MATERIALS,"Type", categories_set_С],
    "CP_Mat_Finish_05": ["02_Mat_Материалы", BuiltInParameterGroup.PG_MATERIALS,"Type", categories_set_С],
    "CP_Mat_Finish_06": ["02_Mat_Материалы", BuiltInParameterGroup.PG_MATERIALS,"Type", categories_set_С],
    "CP_Mat_Finish_07": ["02_Mat_Материалы", BuiltInParameterGroup.PG_MATERIALS,"Type", categories_set_С],
    "CP_Mat_Finish_08": ["02_Mat_Материалы", BuiltInParameterGroup.PG_MATERIALS,"Type", categories_set_С],
    "CP_Mat_Finish_09": ["02_Mat_Материалы", BuiltInParameterGroup.PG_MATERIALS,"Type", categories_set_С],
    "CP_Mat_Finish_10": ["02_Mat_Материалы", BuiltInParameterGroup.PG_MATERIALS,"Type", categories_set_С],
    "CP_Mat_UnitCode": ["02_Mat_Материалы", BuiltInParameterGroup.PG_IDENTITY_DATA,"Type", categories_set_A],
    "CP_Mat_Units": ["02_Mat_Материалы", BuiltInParameterGroup.PG_IDENTITY_DATA,"Type", categories_set_A],
    "CP_Dim_Quantity": ["03_Dim_Размеры", BuiltInParameterGroup.PG_IDENTITY_DATA,"Type", categories_set_A],
    "CP_Sheet_Number": ["04_Sheet_Листы", BuiltInParameterGroup.PG_IDENTITY_DATA,"Instance", [BuiltInCategory.OST_Sheets]],
    "CP_Sheet_Quantity": ["04_Sheet_Листы", BuiltInParameterGroup.PG_IDENTITY_DATA,"Instance", [BuiltInCategory.OST_Sheets]],
    "CP_Sheet_Scale": ["04_Sheet_Листы", BuiltInParameterGroup.PG_IDENTITY_DATA,"Instance", [BuiltInCategory.OST_Sheets]],
    "CP_Sheet_Zone": ["04_Sheet_Листы", BuiltInParameterGroup.PG_IDENTITY_DATA,"Instance", [BuiltInCategory.OST_Sheets]],
    "CP_Sheet_ViewName": ["04_Sheet_Листы", BuiltInParameterGroup.PG_IDENTITY_DATA,"Instance", [BuiltInCategory.OST_Sheets]],
    "CP_View_Number": ["08_View_Виды", BuiltInParameterGroup.PG_IDENTITY_DATA,"Instance", categories_set_D],
    "CP_View_Organization": ["08_View_Виды", BuiltInParameterGroup.PG_IDENTITY_DATA,"Instance", categories_set_D]
    # Добавлять другие параметры здесь
}

# Начинаем транзакцию
with Transaction(doc, "Добавление списка общих параметров из ФОП") as t:
    t.Start()

    for parameter_name, param_info in parameters_to_add.items():
        group_name, parameter_group, instance_or_type, category_list = param_info

        # Поиск группы и параметра в ФОП
        param_definition = None
        for group in shared_params_file.Groups:
            if group.Name == group_name:
                param_definition = group.Definitions.get_Item(parameter_name)
                break

        # Если параметр найден в ФОП
        if param_definition is not None:
            try:
                # Создаем список категорий для привязки
                categories = CategorySet()
                for category_id in category_list:
                    category = doc.Settings.Categories.get_Item(category_id)
                    if category:
                        categories.Insert(category)
                    else:
                        print("Категория {} не найдена в документе.".format(category_id))

                # Создаем привязку параметра к категориям
                if instance_or_type == "Instance":
                    binding = doc.Application.Create.NewInstanceBinding(categories)
                elif instance_or_type == "Type":
                    binding = doc.Application.Create.NewTypeBinding(categories)
                else:
                    print("Неизвестный тип привязки для параметра {}. Пропуск.".format(parameter_name))
                    continue

                # Добавляем параметр в проект
                binding_map = doc.ParameterBindings
                binding_map.Insert(param_definition, binding, parameter_group)
                print("Параметр {} успешно добавлен в проект.".format(parameter_name))
            except Exception as e:
                print("Ошибка при добавлении параметра {} : {}".format(parameter_name, e))
        else:
            # Выводим предупреждение, если параметра нет в ФОП
            print("Параметр {} не найден в группе {} текущего ФОП.".format(parameter_name, group_name))

    # Завершаем транзакцию
    t.Commit()
