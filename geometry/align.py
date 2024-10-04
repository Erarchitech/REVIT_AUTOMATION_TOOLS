# module align.py
# -*- coding: utf-8 -*-
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit import DB

import sys
sys.path.append(r'D:\\Stepik_Python')

class Align(object):
    '''
    Класс-перечисление, позволяющий задавать привязку
        START - привязка по начальной точке
        CENTER - привязка по центру
        END - привязка по конечной точке
    '''
    START = 0
    CENTER = 1
    END = 2