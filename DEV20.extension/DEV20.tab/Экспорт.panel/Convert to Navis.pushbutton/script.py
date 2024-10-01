# -*- coding: utf-8 -*- 
__doc__ = """

"""

__context__ = 'zero-doc'
__title__ = 'Convert to NWC'
__helpurl__ = "https://sites.google.com/rencons.com/mepbimhelp/home/pyrevit/managercp/modelscopy/convert-to-nwc"

from modeltools import *
from Autodesk.Revit.UI import TaskDialog 



def main():
    SelectedModels = SelectModelsForNWC("Select Folder for Models to convert",
        "Select Folder to get Models",
        "Select Models",
        "Select",
        Multiselect = True,
        pathtype = "File")
    # print(SelectedModels)
    ConvertToNWCstream(SelectedModels,1)

if __name__ == '__main__':
    main()
