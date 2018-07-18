#!/usr/bin/python
# -*- coding: utf-8 -*-

from excelreport import ExcelReport, XLSXColor, PrintSet

rep = ExcelReport('Отчёт')

rep.ws['A1'] = 42
for i in range(1, 100):
    rep.ws.append( [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16 ] )

rep.open_excel()
