#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlsreport import *
from xlslabel import *
from xlstableheader import *
from xlstable import *
from sqltabledata import *

import locale
import sys

locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')

THC=XLSTableHeaderColumn
TCI=XLSTableColumnInfo

rep = XLSReport('Акт передачи образцов')

tableheader = XLSTableHeader( columns=(
        THC( 'Артикул',        widths=[20] ),
        THC( 'Цвет ШП/Global', widths=[20] ),
        THC( 'Размеры', struct=[
                THC( 'р', widths=[7] )
            ]*13 ),
        THC( 'Номера коробок', widths=[20] ),
        ) )
max_col = tableheader.column_count
rep.apply_column_widths(tableheader)

cur_row = rep.apply_preamble(max_col)
cur_row = rep.apply_label(XLSLabel('Прибыла ТЕ такого то числа', 1),
                          first_row=cur_row, col_count=max_col)
cur_row = rep.apply_tableheader(tableheader, first_row=cur_row)

table_info = (\
        TCI('ArticleGlobalCode',  'string', 1),
        TCI('OItemColorName',  'string', 1),
        TCI('Sum1',  'int',    1),
        TCI('Sum2',  'int',    1),
        TCI('Sum3',  'int',    1),
        TCI('Sum4',  'int',    1),
        TCI('Sum5',  'int',    1),
        TCI('Sum6',  'int',    1),
        TCI('Sum7',  'int',    1),
        TCI('Sum8',  'int',    1),
        TCI('Sum9',  'int',    1),
        TCI('Sum10', 'int',    1),
        TCI('Sum11', 'int',    1),
        TCI('Sum12', 'int',    1),
        TCI('Sum13', 'int',    1),
        TCI('',      'string', 1),
        )

if sys.platform.startswith('win'):
    sqlprocedure = "ORDERS.dbo.OMS_TRANSPORT_ReportPackageGlobalInvoiceList"
    sqlparamlist = ( "{53DAD87F-8C0F-4178-9A27-9F686E44A8FD}", )
    sqlquery = "EXEC {0:s} {1:s}".format(sqlprocedure, ", ".join(map(maybe_sqlquoted, sqlparamlist)))
    #  print(sqlquery)
    table_data = get_table_data( sqlquery, table_info )
else:
    table_data = ( \
        [ 'MSH05435', 'черный', 50, 0, 150, 100, 200, 0, 200, 0, 0, 0, 0, 0, 0, '1-50'  ],
        [ 'MSH05435', 'белый',   0, 0, 150,   0, 220, 0,   0, 0, 0, 0, 0, 0, 0, '1-50'  ],
        [ 'MSH05436', 'черный', 50, 0, 150, 200, 268, 0,   0, 0, 0, 0, 0, 0, 0, '31-50' ],
        [ 'MSH05437', 'черный',  0, 0, 150, 100, 205, 0, 200, 0, 0, 0, 0, 0, 0, '1-50'  ],
        [ 'MSH05437', 'красный',50, 0, 150, 100, 280, 0,   0, 0, 0, 0, 0, 0, 0, '1-50'  ],
        [ 'MSH05437', 'желтый', 50, 0, 150, 200, 200, 0,   0, 0, 0, 0, 0, 0, 0, '1-50'  ] )

table = XLSTable(table_info, table_data )
cur_row = rep.apply_table(table, first_row=cur_row)

rep.launch_excel()
