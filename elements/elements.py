# -*- coding: utf-8 -*-
import clr
import sys
import operator
sys.path.append(r'D:\Stepik_Python')
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import FilteredElementCollector as FEC
from Custom_Functions import *

doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = uiapp.ActiveUIDocument


def get_type_name(element):
    if hasattr(element, 'GetTypeId'):
        element_type = doc.GetElement(element.GetTypeId())
        return element_type.Paramter[DB.BuiltInParameter.SYMBOL_NAME_PARAM].AsString()


def is_of_type(element, type_name):
    return type_name == get_type_name(element)