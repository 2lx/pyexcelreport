#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlsreport import *

rep = XLSReport('Отчёт')

col_count = rep.make_tableheader_1line( headers=(\
    [ 'Артикул' ], \
    [ 'Цвет ШП/Global', 1 ], \
    [ 'Размеры', 13 ], \
    [ 'Номера коробок', 1 ] \
    ), first_row=4, first_col=2 )

rep.launch_excel()
