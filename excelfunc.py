#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
import tempfile
from openpyxl import Workbook, worksheet

import excelreport

def workbook_create():
    wb = Workbook()
    for i in wb.worksheets:
        wb.remove(i)
    return wb

def sheet_create(wb, main_sheet_name):
    wb.create_sheet( main_sheet_name )
    ws = wb.active

    return ws

def sheet_print_setup(ws, print_settings):
    'https://openpyxl.readthedocs.io/en/2.5/_modules/openpyxl/worksheet/page.html'

    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = False
    ws.print_options.headings = False
    ws.print_options.gridLines = False

    ws.page_margins.left = 0.2
    ws.page_margins.right = 0.2
    ws.page_margins.top = 0.2
    ws.page_margins.bottom = 0.2
    ws.page_margins.header = 0
    ws.page_margins.footer = 0

    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToHeight = False

    if print_settings in [ excelreport.PrintSet.LandscapeW1, excelreport.PrintSet.LandscapeW2 ]:
        worksheet.Worksheet.set_printer_settings(ws, paper_size = 1, orientation='landscape')

    if print_settings in [ excelreport.PrintSet.PortraitW1 ]:
        worksheet.Worksheet.set_printer_settings(ws, paper_size = 1, orientation='portrait')

    if print_settings in [ excelreport.PrintSet.LandscapeW1, excelreport.PrintSet.PortraitW1 ]:
        ws.page_setup.fitToWidth = 1

    if print_settings in [ excelreport.PrintSet.LandscapeW2 ]:
        ws.print_options.horizontalCentered = False
        ws.page_setup.fitToWidth = 2

def temporary_file(filename='sample'):
    cnt = 1
    while True:
        newfilename = '{0:s}\\{1:s}{2:0>2d}.xlsx'.format( tempfile.gettempdir(), filename, cnt )
        if os.path.isfile(newfilename):
            cnt+=1
        else:
            break

    return newfilename
