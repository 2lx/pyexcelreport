#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
from enum import Enum

import excelfunc

class PrintSet(Enum):
    PortraitW1 = 1
    LandscapeW1 = 3
    LandscapeW2 = 4

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

	def launch_excel( self, filename='sample' ):
		''
		newfilename = excelfunc.temporary_file( filename )
		self.wb.save( newfilename )

		print( "Opening file '{0:s}'...".format( newfilename ) )
		os.system( 'start excel.exe {0:s}'.format( newfilename ) )

	def make_tableheader_1line( self, headers, first_row=1, first_col=1 ):
		''
		self.ws.row_dimensions[ first_row ].height = 50

		cur_col = first_col
		for titleopt in headers:
			title = titleopt[0]

			cnt = 1
			if len(titleopt) > 1:
				cnt = titleopt[1]
				self.ws.merge_cells( start_row=first_row, start_column=cur_col, \
								end_row=first_row, end_column=cur_col + cnt - 1 )

			self.ws.cell( row=first_row, column=cur_col, value=title )
			cur_col+=cnt

		excelfunc.apply_border( self.ws, first_row, first_col, first_row, cur_col - 1 )
		excelfunc.apply_outline( self.ws, first_row, first_col, first_row, cur_col - 1, 'medium' )

		return len(headers)
