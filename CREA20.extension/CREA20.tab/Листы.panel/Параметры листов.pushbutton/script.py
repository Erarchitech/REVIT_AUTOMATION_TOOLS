# -*- coding: utf-8 -*-
import clr
import sys
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import FilteredElementCollector as FEC

uiapp = __revit__ # noqa
app = __revit__.Application # noqa
uidoc = __revit__.ActiveUIDocument # noqa
doc = __revit__.ActiveUIDocument.Document # noqa

Map = doc.ParameterBindings
parameter = Map.ForwardIterator()

def get_parameter_value(parameter):
    if isinstance(parameter, Parameter):
        storage_type = parameter.StorageType
        if storage_type:
            exec("parameter_value = parameter.As{}()".format(storage_type))
            return parameter_value

sheets = FEC(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements()
Dict_stage = {
	"А":["CP_Sheet_Stage","АГО"],
	"Э":["CP_Sheet_Stage","ЭП"],
	"П":["CP_Sheet_Stage","П"],
	"К":["CP_Sheet_Stage","КРД"]
}
Dict_Count = {}
check = []

with Transaction(doc, 'Fill Sheet Parameters') as t:
	t.Start()
	for sheet in sheets:
		number = get_parameter_value(sheet.Parameter[BuiltInParameter.SHEET_NUMBER])
		stage = number[:1]
		if stage in Dict_stage.keys():
			zone = number[1:2]
			view_organization = zone + "_" + Dict_stage[stage][1]
			number_sheet = number[2:]
			sheet.LookupParameter(Dict_stage[stage][0]).Set(Dict_stage[stage][1])
			sheet.LookupParameter("CP_Sheet_Zone").Set(zone)
			sheet.LookupParameter("CP_Sheet_Number").Set(number_sheet)
			sheet.LookupParameter("CP_View_Organization").Set(view_organization)
			if get_parameter_value(sheet.Parameter[BuiltInParameter.SHEET_SCHEDULED]) == 1:
				if str(number[:2]) in Dict_Count.keys():
					Dict_Count[str(number[:2])] += 1
				else:
					Dict_Count[str(number[:2])] = 1
	for sheet in sheets:
		if get_parameter_value(sheet.Parameter[BuiltInParameter.SHEET_SCHEDULED]) == 1:
			number = get_parameter_value(sheet.Parameter[BuiltInParameter.SHEET_NUMBER])
			stage = number[:1]
			if stage in Dict_stage.keys():
				key = str(number[:2])
				try:
					Count = str(Dict_Count[key])
					check.append(Count)
					sheet.LookupParameter("CP_Sheet_Quantity").Set(Count)
				except:
					pass
	t.Commit()
print(check)
print("Параметры CP_Sheet_Stage, CP_Sheet_Zone, CP_Sheet_Number успешно записаны")