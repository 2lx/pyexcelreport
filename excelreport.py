#!/usr/bin/python
# -*- coding: utf-8 -*-

import tempfile
import os
from enum import Enum

import excelfunc

class PrintSet(Enum):
    PortraitW1 = 1
    LandscapeW1 = 3
    LandscapeW2 = 4

class XLSXColor(Enum):
	colorLightGray = 0xCCCCCC

class ExcelReport():
	'Класс, инкапсулирующий в себе методы для создания отчета в Excel'

	def __init__( self, main_sheet_name='Новый лист', print_settings=PrintSet.LandscapeW1 ):
		'Конструктор, принимает название страницы и параметры печати'
		self.wb = excelfunc.workbook_create()
		self.ws = excelfunc.sheet_create( self.wb, main_sheet_name )
		excelfunc.sheet_print_setup( self.ws, print_settings )

	def open_excel( self, filename='sample' ):
		''
		newfilename = excelfunc.temporary_file( filename )
		self.wb.save( newfilename )

		print( "Opening file '{0:s}'...".format( newfilename ) )
		os.system( 'start excel.exe {0:s}'.format( newfilename ) )
