# -*- coding: utf-8 -*-
import clr
import sys
sys.path.append(r'D:\\Stepik_Python')
clr.AddReference('RevitServices')
clr.AddReference('RevitNodes')
from RevitServices.Persistence import DocumentManager
from Autodesk.Revit.DB import FilteredElementCollector as FEC
from System.Collections.Generic import List
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import Selection as SEL
from Custom_Functions import *

doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = uiapp.ActiveUIDocument

class CategoriesSelectionFilter(SEL.ISelectionFilter):
        def __init__(self, b_categories):
            super(CategoriesSelectionFilter, self).__init__()
            self._category_ids = [DB.ElementId(b_category)
                                for b_category in to_list(b_categories)]

        def AllowElement(self, element):
            return element.Category.Id in self._category_ids

class LevelCategoriesSelectionFilter(SEL.ISelectionFilter):
        def __init__(self, dict_level_categories = {"level_name": [DB.BuiltInCategory.OST_Walls]}):
            super(LevelCategoriesSelectionFilter, self).__init__()
            dict_converted = {}
            #конвертируем словарь в {Level_id:[b_category_ids]}
            for level_name in dict_level_categories.keys():
                level_id = get_element_by_name(doc, level_name, DB.Level).Id.IntegerValue
                category_ids = [DB.ElementId(b_category).IntegerValue \
                    for b_category in to_list(dict_level_categories[level_name])]
                dict_converted[level_id] = category_ids
            self._level_ids = [key for key in dict_converted.keys()]
            self._category_ids = [value for value in dict_converted.values()]
            self._levelcategories_ids = dict_converted
                
        def AllowElement(self, element):
            level_id = element.LevelId.IntegerValue
            category_id = element.Category.Id.IntegerValue
            element_ids = []
            if level_id in self._levelcategories_ids.keys():
                if category_id in self._levelcategories_ids[level_id]:
                    element_ids.append(element.Id.IntegerValue)
            return uidoc.Selection.SetElementIds(element_ids)