# module unit.py
# -*- coding: utf-8 -*-
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit import DB

# Получаем версию Revit
app_version = __revit__.Application.VersionNumber


class Unit(object):
    def __init__(self, doc, value, is_display=True, unit_type=None):
        self._doc = doc
        self._unit_type = self._get_unit_type(unit_type)
        self._internal_value = self._convert_value(value, is_display) if is_display else value

    def _get_unit_type(self, unit_type):
        """Определяет тип единиц в зависимости от версии Revit"""
        if int(app_version) < 2022:
            # Используем старый тип для версий Revit 2021 и младше
            return DB.DisplayUnitType.DUT_MILLIMETERS if unit_type is None else unit_type
        else:
            # Используем новый тип для Revit 2022 и новее
            return DB.UnitTypeId.Millimeters if unit_type is None else unit_type

    def _convert_value(self, value, to_internal):
        """Конвертация значения с проверкой измеряемости типа"""
        conversion_method = DB.UnitUtils.ConvertToInternalUnits if to_internal else DB.UnitUtils.ConvertFromInternalUnits
        return conversion_method(value, self.unit_type)

    @property
    def unit_type(self):
        """Тип единиц (UnitType или DisplayUnitType в зависимости от версии)"""
        return self._unit_type

    @property
    def display_units(self):
        """Единицы измерения, отображаемые в интерфейсе Revit (DisplayUnitType)"""
        format_options = self._doc.GetUnits().GetFormatOptions(self._unit_type)
        return format_options.DisplayUnits

    @property
    def internal(self):
        """Числовое значение во внутренних единицах Revit"""
        return self._internal_value

    @property
    def display(self):
        """Числовое значение в единицах интерфейса Revit"""
        return self._convert_value(self._internal_value, False)