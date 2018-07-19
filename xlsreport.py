#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
from enum import Enum

from xlsfunc import *
from difffunc import *

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

    def __init__( self, main_sheet_name='Новый лист', print_conf=PrintConf.LandscapeW1 ):
        """Конструктор, принимает название страницы и параметры печати
        """
        self.wb = workbook_create()
        self.ws = sheet_create( self.wb, main_sheet_name )
        sheet_print_setup( self.ws, print_conf.value[0], print_conf.value[1] )

    def launch_excel(self, templatename='sample' ):
        """Запускает программу по умолчанию для xls-файлов и открывает в ней workbook
        """
        newfilename = temporary_file( templatename )
        self.wb.save( newfilename )

        print( "Opening file '{0:s}'...".format( newfilename ) )
        open_file( newfilename )

    def apply_table(self, table, first_row=1, first_col=1):
        table.apply_header(self.wb.active, first_row, first_col)
