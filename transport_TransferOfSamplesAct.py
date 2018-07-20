#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlsreport import *
from xlstableheader import *

THC=XLSTableHeaderColumn

rep = XLSReport('Акт передачи образцов')

tableheader = XLSTableHeader( headers=(\
    THC( 'Артикул' ), \
    THC( 'Цвет ШП/Global' ), \
    THC( 'Размеры', 13 ), \
    THC( 'Номера коробок' ) \
    ) )
max_col = tableheader.column_count()
print(max_col)

rep.apply_tableheader(tableheader, first_row=4)
rep.apply_tableheader(tableheader, first_row=6)

print("OK")
rep.launch_excel()
