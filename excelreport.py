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

class XLSXReport():
	'Класс, инкапсулирующий в себе методы для создания отчета в Excel'

	def __init__( self, main_sheet_name='Новый лист', print_settings=PrintSet.LandscapeW1 ):
		'Конструктор, принимает название страницы и параметры печати'
		self.wb = excelfunc.workbook_create()
		self.ws = excelfunc.sheet_create( self.wb, main_sheet_name, print_settings )

		self.ws['A1'] = 42
		for i in range(1, 100):
			self.ws.append( [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16 ] )

	def open_excel( self, filename='sample' ):
		''

		cnt = 1
		while True:
			newfilename = '{0:s}\\{1:s}{2:0>2d}.xlsx'.format( tempfile.gettempdir(), filename, cnt )
			if os.path.isfile(newfilename):
				cnt+=1
			else:
				break

		self.wb.save( newfilename )
		print( "Opening file '{0:s}'...".format( newfilename ) )
		os.system( 'start excel.exe {0:s}'.format( newfilename ) )
