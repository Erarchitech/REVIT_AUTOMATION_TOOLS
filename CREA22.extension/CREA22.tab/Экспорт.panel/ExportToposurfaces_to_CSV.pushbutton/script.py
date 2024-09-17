# -*- coding: utf-8 -*-
import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('RevitServices')
# sys.path.append(r'C:\Users\Евгений\Documents\Stepik_Python\Custom_Functions.py')

from RevitServices.Persistence import DocumentManager

from Autodesk.Revit.DB import *
import csv

import pyrevit
from pyrevit import forms

uiapp = __revit__ # noqa
app = __revit__.Application # noqa
uidoc = __revit__.ActiveUIDocument # noqa
doc = __revit__.ActiveUIDocument.Document # noqa

def flatten(element, flat_list=None):
    if flat_list is None:
        flat_list = []
    if hasattr(element, "__iter__"):
        for item in element:
            flatten(item, flat_list)
    else:
        flat_list.append(element) 
    return flat_list

# Find the topography surface (you may need to adjust this filter based on your project)
topography_surfaces = FilteredElementCollector(doc).OfCategory(
    BuiltInCategory.OST_Topography
).WhereElementIsNotElementType().ToElements()

topography_points = []

for topo in topography_surfaces:
    if topo:
        # Get the points of the topography surface
        topography_points.append(topo.GetPoints())
topography_points = flatten(topography_points)

# Define a CSV file path to save the points
folder_path = forms.pick_folder(title=None, owner=None)
csv_file_path = folder_path + "\TopographyPoints.csv"
        
# Write the points to a CSV file
with open(csv_file_path, mode="w") as file:
    writer = csv.writer(file)
    writer.writerow(["X", "Y", "Z"])
    for point in topography_points:
        writer.writerow([point.X, point.Y, point.Z])
print("CSV file is written\n")
print(topography_points)
