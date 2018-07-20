#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlsreport import *
from xlslabel import *
from xlstableheader import *
from xlstable import *

THC=XLSTableHeaderColumn
TCI=XLSTableColumnInfo

rep = XLSReport('Акт передачи образцов')

tableheader = XLSTableHeader( headers=(\
    THC( 'Артикул',         width=20 ), \
    THC( 'Цвет ШП/Global',  width=20 ), \
    THC( 'Размеры',         struct=[\
        THC( title='размер', width=7 ) ]*13 \
        ), \
    THC( 'Номера коробок',  width=20 ) \
    ) )
max_col = tableheader.column_count()
print(max_col)

cur_row = rep.apply_preamble(max_col)
cur_row = rep.apply_label(XLSLabel('Прибыла ТЕ такого то числа', 1), first_row=cur_row, col_count=max_col)
cur_row = rep.apply_tableheader(tableheader, first_row=cur_row)

table_info = (\
    TCI('field1',  'string', 1), \
    TCI('field2',  'string', 1), \
    TCI('field3',  'int',    1), \
    TCI('field4',  'int',    1), \
    TCI('field5',  'int',    1), \
    TCI('field6',  'int',    1), \
    TCI('field7',  'int',    1), \
    TCI('field8',  'int',    1), \
    TCI('field9',  'int',    1), \
    TCI('field10', 'int',    1), \
    TCI('field11', 'int',    1), \
    TCI('field12', 'int',    1), \
    TCI('field13', 'int',    1), \
    TCI('field14', 'int',    1), \
    TCI('field15', 'int',    1), \
    TCI('field16', 'string', 1), \
    )

table_data = ( \
        [ 'MSH05435', 'черный', 50, 0, 150, 100, 200, 0, 200, 0, 0, 0, 0, 0, 0, '1-50'  ],
        [ 'MSH05435', 'белый',   0, 0, 150,   0, 220, 0,   0, 0, 0, 0, 0, 0, 0, '1-50'  ],
        [ 'MSH05435', 'розовый', 0, 0,   0, 150, 120, 0, 200, 0, 0, 0, 0, 0, 0, '7-30'  ],
        [ 'MSH05436', 'черный', 50, 0, 150, 200, 268, 0,   0, 0, 0, 0, 0, 0, 0, '31-50' ],
        [ 'MSH05437', 'черный',  0, 0, 150, 100, 205, 0, 200, 0, 0, 0, 0, 0, 0, '1-50'  ],
        [ 'MSH05437', 'красный',50, 0, 150, 100, 280, 0,   0, 0, 0, 0, 0, 0, 0, '1-50'  ],
        [ 'MSH05437', 'желтый', 50, 0, 150, 200, 200, 0,   0, 0, 0, 0, 0, 0, 0, '1-50'  ] )

table = XLSTable(table_info, table_data )

cur_row = rep.apply_table(table, first_row=cur_row)

rep.launch_excel()
print("OK")
