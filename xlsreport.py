#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
from enum import Enum

from xlsutils import *
from systemutils import *

class PrintConf(Enum):
    """Настройки печати
    """
    PortraitW1  = ('portrait',  1)
    LandscapeW1 = ('landscape', 1)
    LandscapeW2 = ('landscape', 2)

class XLSColor(Enum):
    """Цвета в отчёте
    """
    lGray = 0xCCCCCC

class XLSReport():
    """Класс, инкапсулирующий в себе методы для создания отчета в Excel
    """

    def __init__(self, sheet_name='Новый лист', print_conf=PrintConf.LandscapeW1):
        """Конструктор, создает книгу с одним именованным листом, устанавливает параметры для печати
        """
        self._wb = workbook_create()
        self._ws = sheet_create(self._wb, sheet_name)
        sheet_print_setup(self._ws, print_conf.value[0], print_conf.value[1])

    def launch_excel(self, templatename='sample'):
        """Запускает программу по умолчанию для xls-файлов и открывает в ней workbook
        """
        newfilename = temporary_file(templatename)
        self._wb.save(newfilename)

        print("Opening file '{0:s}'...".format(newfilename))
        open_file(newfilename)

    def apply_tableheader(self, tableheader, first_row=1, first_col=1):
        tableheader.apply(self._wb.active, first_row, first_col)
