#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlsreport import *
from xlstable import *

HDR=XLSTableHeader

rep = XLSReport('Акт передачи образцов')

table = XLSTable( headers=(\
    HDR( 'Артикул' ), \
    HDR( 'Цвет ШП/Global' ), \
    HDR( 'Размеры', 13 ), \
    HDR( 'Номера коробок' ) \
    ) )
max_col = table.column_count()
print(max_col)
rep.apply_table(table, first_row=4)

print("OK")
rep.launch_excel()
