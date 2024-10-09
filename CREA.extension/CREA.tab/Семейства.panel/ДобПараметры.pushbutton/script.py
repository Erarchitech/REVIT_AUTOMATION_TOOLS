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

def get_external_definition(app,group_name,definition_name):
    return app.OpenSharedParameterFile().Groups[group_name].Definitions[definition_name]

Parameter_dict = {
    "CP_Gen_Mark": ["01_Gen_Общие", BuiltInParameterGroup.PG_TEXT],
    "CP_Gen_Description": ["01_Gen_Общие", BuiltInParameterGroup.PG_TEXT],
    "CP_Gen_Name": ["01_Gen_Общие", BuiltInParameterGroup.PG_TEXT],
    "CP_Dim_Height": ["03_Dim_Размеры", BuiltInParameterGroup.PG_GEOMETRY],
    "CP_Dim_Length": ["03_Dim_Размеры", BuiltInParameterGroup.PG_GEOMETRY],
    "CP_Dim_Width": ["03_Dim_Размеры", BuiltInParameterGroup.PG_GEOMETRY],
    "CP_Gen_Image": ["01_Gen_Общие", BuiltInParameterGroup.PG_GRAPHICS],
    "CP_Gen_SurfacePattern": ["01_Gen_Общие", BuiltInParameterGroup.PG_GRAPHICS],
    "CP_Gen_Ver": ["01_Gen_Общие", BuiltInParameterGroup.PG_IDENTITY_DATA],
}

try:
    f_manager = doc.FamilyManager
except Exception as e:
    print("ОШИБКА", str(e))
    sys.exit()

def create_family_parameter(f_manager, external_definition, p_group=BuiltInParameterGroup.INVALID, Inst_parameter=False):
    try:
        f_manager.AddParameter(external_definition, p_group, Inst_parameter)
        message = str(external_definition.Name + " - параметр успешно добавлен" )
    except:
        message = str(external_definition.Name + " - параметр уже существует" )
    return message

with Transaction(doc, 'Добавлены CP параметры') as t:
	t.Start()
	for key, value in Parameter_dict.items():
		ext_def = get_external_definition(app, value[0], key)
		print str(create_family_parameter(f_manager, ext_def, p_group=value[1], Inst_parameter=False))
	t.Commit()


