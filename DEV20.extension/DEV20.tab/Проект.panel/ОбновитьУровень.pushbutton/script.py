# -*- coding: utf-8 -*-
import clr
clr.AddReference('RevitServices')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System')
clr.AddReference('System.Windows')
clr.AddReference('System.Windows.Forms')
clr.AddReference('IronPython.Wpf')

from System import Array
from System.Collections.Generic import List
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import FilteredElementCollector, ElementLevelFilter, ElementTransformUtils, BuiltInParameter, Level, LocationPoint, LocationCurve, XYZ, Transaction
from RevitServices.Persistence import DocumentManager
import pyrevit
from pyrevit import forms
from pyrevit import script
from pyrevit.forms import WPFWindow

# Get active document
uiapp = __revit__ # noqa
app = __revit__.Application # noqa
uidoc = __revit__.ActiveUIDocument # noqa
doc = __revit__.ActiveUIDocument.Document # noqa

collector_levels = FilteredElementCollector(doc).OfClass(Level)

level_list = []
for lev in collector_levels:
    level_list.append(lev.Name)
# print(level_list)

old_level_name = pyrevit.forms.SelectFromList.show(context = level_list, title = 'Old Level', width = 700, height = 300,button_name = 'Select Old Level')
new_level_name = pyrevit.forms.SelectFromList.show(context = level_list, title = 'Old Level', width = 700, height = 300,button_name = 'Select New Level')

# If levels are selected, proceed with reassignment logic
if old_level_name and new_level_name:
    # Function to get level by name
    def get_level_by_name(level_name):
        collector = FilteredElementCollector(doc).OfClass(Level)
        for level in collector:
            if level.Name == level_name:
                return level
        return None

    # Function to get elements on a specific level
    def get_elements_on_level(level_id):
        level_filter = ElementLevelFilter(level_id)
        collector = FilteredElementCollector(doc)
        elements = collector.WherePasses(level_filter).WhereElementIsNotElementType().ToElements()
        return elements
    
    def calculate_elevation_difference(level1, level2):
        if level1 and level2:
            elevation1 = UnitUtils.ConvertFromInternalUnits(level1.Elevation,DisplayUnitType.DUT_MILLIMETERS)
            elevation2 = UnitUtils.ConvertFromInternalUnits(level2.Elevation,DisplayUnitType.DUT_MILLIMETERS)
            return (elevation1 - elevation2)
        else:
            raise ValueError("One or both levels not found.")
        
    def is_loaded_family(element):
        try:
            family = element.Symbol.Family
            # Check if the family is loaded (editable) and not an in-place family
            try:
                result = family.IsEditable and not family.IsInPlace
                if result:
                    if not element.SuperComponent:
                        return result
                    print('Семейство экземпляра является вложенным: {}'.format(element.Id))
            except:
                print("Ошибка в индентификации загружаемого семейства: " + '{}'.format(element.Id))
        except:
                return False

    # Function to set a new level and maintain position with a new offset
    def reassign_level_and_offset(doc, elements, old_level, new_level, Categories_Dict):
        with Transaction(doc, 'Изменить уровень элементов') as t:
            t.Start()
            counter = 0
            for element in elements:
                # print(element.Symbol.Family.Name)
                # print(element.Id)
                # try:
                if is_loaded_family(element):
                    # Update the level parameter
                    level_param = element.Parameter[BuiltInParameter.FAMILY_LEVEL_PARAM]
                    offset_param = element.Parameter[BuiltInParameter.INSTANCE_ELEVATION_PARAM]
                    offset_value = UnitUtils.ConvertFromInternalUnits(offset_param.AsDouble(),DisplayUnitType.DUT_MILLIMETERS) + calculate_elevation_difference(old_level, new_level)
                    # print(offset_value)
                    if level_param and offset_param:
                        try:
                            level_param.Set(new_level.Id)
                            offset_param.Set(UnitUtils.ConvertToInternalUnits(offset_value,DisplayUnitType.DUT_MILLIMETERS))
                            print('Параметры успешно записаны')
                            counter +=1
                        except:
                            print('Элемент не перемещен или перемещен некорректно: {}'.format(element.Id))
                            pass
                    else:
                        print('Параметры не обнаружены')
                else:
                    # print('check_families_walls')
                    wall_category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Walls)
                    if element.Category.Id == wall_category.Id:
                        level_param = element.Parameter[Categories_Dict['OST_Walls'][0]]
                        top_constraint_param = element.Parameter[BuiltInParameter.WALL_HEIGHT_TYPE]
                        top_offset_param = element.Parameter[BuiltInParameter.WALL_TOP_OFFSET]
                        offset_param = element.Parameter[Categories_Dict['OST_Walls'][1]]
                        height_param = element.Parameter[Categories_Dict['OST_Walls'][2]]
                        offset_value = UnitUtils.ConvertFromInternalUnits(offset_param.AsDouble(),DisplayUnitType.DUT_MILLIMETERS) + calculate_elevation_difference(doc.GetElement(level_param.AsElementId()), new_level)
                        top_offset_value = UnitUtils.ConvertFromInternalUnits(top_offset_param.AsDouble(),DisplayUnitType.DUT_MILLIMETERS) + calculate_elevation_difference(doc.GetElement(top_constraint_param.AsElementId()), new_level)
                        if level_param and offset_param:
                            try:
                                offset_param.Set(UnitUtils.ConvertToInternalUnits(offset_value,DisplayUnitType.DUT_MILLIMETERS))
                                level_param.Set(new_level.Id)
                                top_constraint_param.Set(new_level.Id)
                                top_offset_param.Set(UnitUtils.ConvertToInternalUnits(top_offset_value,DisplayUnitType.DUT_MILLIMETERS))
                                print('Параметры успешно записаны')
                                counter +=1
                            except:
                                pass
                        else:
                            print('Параметры не обнаружены')
                    else:
                        # print('check_families_not_walls')
                        for cat in Categories_Dict.keys():
                            category_enum_value = getattr(BuiltInCategory, cat)
                            dict_category = doc.Settings.Categories.get_Item(category_enum_value)
                            if element.Category.Id == dict_category.Id:
                                # print(Categories_Dict[cat][0])
                                level_param = element.Parameter[Categories_Dict[cat][0]]
                                offset_param = element.Parameter[Categories_Dict[cat][1]]
                                offset_value = UnitUtils.ConvertFromInternalUnits(offset_param.AsDouble(),DisplayUnitType.DUT_MILLIMETERS) + calculate_elevation_difference(old_level, new_level)
                                if level_param and offset_param:
                                    level_param.Set(new_level.Id)
                                    try:
                                        offset_param.Set(UnitUtils.ConvertToInternalUnits(offset_value,DisplayUnitType.DUT_MILLIMETERS))
                                        print('Параметры успешно записаны')
                                        counter +=1
                                    except:
                                        pass
                                else:
                                    print('Параметры не обнаружены')
                # except:
                #     print("Ошибка при обновлении семейства: " + '{}'.format(element.Id))
                #     pass
                # print('_______')
            t.Commit()
        return counter


    Categories_Dict = {
        #Стены: [Зависимость снизу, Смещение снизу, Неприсоединенная высота]
        'OST_Walls': [BuiltInParameter.WALL_BASE_CONSTRAINT, BuiltInParameter.WALL_BASE_OFFSET, BuiltInParameter.WALL_USER_HEIGHT_PARAM],
        #Перекрытия: [Уровень, Смещение от уровня]
        'OST_Floors': [BuiltInParameter.LEVEL_PARAM, BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM],
        #Потолки: [Уровень, Смещение от уровня]
        'OST_Ceilings': [BuiltInParameter.LEVEL_PARAM, BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM],
        #Осветительные приборы: Уровень
    }


    # Get the selected levels by name
    old_level = get_level_by_name(old_level_name)
    new_level = get_level_by_name(new_level_name)

    if old_level and new_level:
        # Collect elements from the old level
        elements_on_old_level = get_elements_on_level(old_level.Id)

        # Reassign elements to new level with a new offset (for example, 1 foot)
        new_offset = calculate_elevation_difference(old_level, new_level)
        # print(new_offset)
        elements_on_new_level = reassign_level_and_offset(doc, elements_on_old_level, old_level, new_level, Categories_Dict)
        print('Общее количество элементов на исходном уровне: {}'.format(len(elements_on_old_level)))
        print('Общее количество элементов на целевом уровне после переопределения: {}'.format(elements_on_new_level))
    else:
        if not old_level:
            print("Уровень '{old_level_name}' не найден.")
        if not new_level:
            print("Еровень '{new_level_name}' не найден.")
else:
    print("Пожалуйста выберите оба уровня - с какого и на какой необходимо переназначить элементы.")
