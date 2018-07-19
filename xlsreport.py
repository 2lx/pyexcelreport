#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
from enum import Enum

from xlsfunc import *
from difffunc import *

class PrintSet(Enum):
    PortraitW1 = 1
    LandscapeW1 = 3
    LandscapeW2 = 4

class XLSXColor(Enum):
    colorLightGray = 0xCCCCCC

class XLSReport():
    'Класс, инкапсулирующий в себе методы для создания отчета в Excel'

    def __init__( self, main_sheet_name='Новый лист', print_settings=PrintSet.LandscapeW1 ):
        'Конструктор, принимает название страницы и параметры печати'
        self.wb = workbook_create()
        self.ws = sheet_create( self.wb, main_sheet_name )
        porientation = 'portrait' if print_settings in [ PrintSet.PortraitW1 ] else 'landscape'
        pwidth = 2 if print_settings in [ PrintSet.LandscapeW2 ] else 1
        sheet_print_setup( self.ws, porientation, pwidth )

    def launch_excel( self, templatename='sample' ):
        'Запускает программу по умолчанию для xls-файлов и открывает в ней workbook'
        newfilename = temporary_file( templatename )
        self.wb.save( newfilename )

        print( "Opening file '{0:s}'...".format( newfilename ) )
        #  os.system( 'start excel.exe {0:s}'.format( newfilename ) )
        open_file( newfilename )

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

        apply_border( self.ws, first_row, first_col, first_row, cur_col - 1 )
        apply_outline( self.ws, first_row, first_col, first_row, cur_col - 1, 'medium' )

        return len(headers)
