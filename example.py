#!/usr/bin/python
# -*- coding: utf-8 -*-

import locale
import sys

locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')

from xlsreport import *

THC=XLSTableHeaderColumn
TCI=XLSTableColumnInfo

rep = XLSReport('Акт передачи образцов')

# задаю структуру шапки отчёта
tableheader = XLSTableHeader( columns=(
        THC( 'Артикул',         widths=[20] ),
        THC( 'Цвет ШП/Global',  widths=[20] ),
        THC( 'Размеры, составной заголовок', struct=[ THC('р', widths=[ 7]) ]*13 ),
        #  THC( 'Размеры',         widths=[ 7]*13 ),
        THC( 'Номера коробок',  widths=[20] ),
        ) )

max_col = tableheader.column_count
rep.apply_column_widths(tableheader)

# добавляю информациюю о сгенерированном отчёте
cur_row = rep.apply_preamble(max_col)

# добавляю заголовок отчёта
cur_row = rep.apply_label(XLSLabel('Заявка. Прибыла ТЕ такого то числа, а ещё прибыла Партия', LabelHeading.h1),
                          first_row=cur_row, col_count=max_col)
# вывожу доп. информационные поля
rep.apply_label(XLSLabel('От кого: склад "Распределение"', LabelHeading.h5),
                          first_row=cur_row, col_count=6)
cur_row = rep.apply_label(XLSLabel('Кому: склад "ШП технологи"', LabelHeading.h5),
                          first_row=cur_row, first_col=7, col_count=max_col - 7 + 1)

# печатаю шапку отчёта
cur_row = rep.apply_tableheader(tableheader, first_row=cur_row)

# задаю структуру контекста отчёта
table_info = (\
        TCI('ArticleGlobalCode',    'string', 1),
        TCI('OItemColorName',       'string', 1),
        TCI('Sum1',                 'int',    1),
        TCI('Sum2',                 'int',    1),
        TCI('Sum3',                 'int',    1),
        TCI('Sum4',                 'int',    1),
        TCI('Sum5',                 'int',    1),
        TCI('Sum6',                 'int',    1),
        TCI('Sum7',                 'int',    1),
        TCI('Sum8',                 'int',    1),
        TCI('Sum9',                 'int',    1),
        TCI('Sum10',                'int',    1),
        TCI('Sum11',                'int',    1),
        TCI('Sum12',                'int',    1),
        TCI('Sum13',                'int',    1),
        TCI('',                     'string', 1),
        )

# указываю данные для контекста
if sys.platform.startswith('win'):
#    sqlprocedure = "ORDERS.dbo.OMS_TRANSPORT_ReportPackageGlobalInvoiceList"
#    sqlparamlist = ( "{53DAD87F-8C0F-4178-9A27-9F686E44A8FD}", )
#    sqlquery = "EXEC {0:s} {1:s}".format(sqlprocedure, ", ".join(map(maybe_sqlquoted, sqlparamlist)))
#    #  print(sqlquery)
#    table_data = get_mssql_data(sqlquery, table_info)
#else:
    table_data = (
        [ 'MSH05435', 'черный', 50, 0, 150, 100, 200, 0, 200, 0, 0, 0, 0, 0,   0,  '1-50' ],
        [ 'MSH05435', 'черный', 50, 0, 150, 200, 268, 0,   0, 0, 0, 0, 0, 0,   0, '31-50' ],
        [ 'MSH05436', 'белый',   0, 0, 150,   0, 220, 0,   0, 0, 0, 0, 0, 0,   0,  '1-50' ],
        [ 'MSH05437', 'белый',   0, 0, 150, 100, 205, 0, 200, 0, 0, 0, 0, 0, 430,  '1-50' ],
        [ 'MSH05437', 'черный',  0, 0, 150, 100, 205, 0, 200, 0, 0, 0, 0, 0,   0,  '1-50' ],
        [ 'MSH05437', 'черный',  0, 0, 150, 100, 205, 0, 200, 0, 0, 0, 0, 0,   0,  '1-50' ],
        [ 'MSH05437', 'красный',50, 0, 150, 100, 280, 0,   0, 0, 0, 0, 0, 0,   0,  '1-50' ],
        [ 'MSH05437', 'желтый', 50, 0, 150, 200, 200, 0,   0, 0, 0, 0, 0, 0,   0,  '1-50' ] )

table = XLSTable(table_info, table_data)

# указываю столбцы, которые можно скрыть если все значения в контексте нулевые
fn = lambda x: x == 0
for i in range(1, 14):
    table.add_hide_column_condition("Sum{0:d}".format(i), fn)

# определяю функции для отрисовки подзаголовков таблицы
def my_header_func(ws, cur_row, first_col):
    # в функции подзаголовка можно рисовать как обычно, и в том числе вызывать методы report
    colheaders = [ THC("р. {0:d}".format(i)) for i in range(1, 14) ]
    my_subtitle_header = XLSTableHeader( columns=[
            THC('Составной подзаголовок модели', struct=colheaders)],
            row_height=16 )

    cur_row = rep.apply_tableheader(my_subtitle_header, first_row=cur_row + 1, first_col=first_col + 2)
    return cur_row

def my_second_header(ws, cur_row, first_col):
    # просто вывожу надпись
    return rep.apply_label(XLSLabel('Подзаголовок цветомодели', LabelHeading.h3),
                          first_row=cur_row, first_col=first_col + 2, col_count=13)

# указываю поля, которые участвуют в группировках равных значений,
# указываю поля с подитогами, указываю поля с подзаголовками
# всё это указываю в порядке приоритета
table.add_hierarchy_field('ArticleGlobalCode', merging=True,
        subtotal=['Sum1', 'Sum2', 'Sum3', 'Sum4', 'Sum5', 'Sum6', 'Sum7', 'Sum8', 'Sum9', 'Sum10', 'Sum11', 'Sum12', 'Sum13'],
        subtitle=my_header_func)
table.add_hierarchy_field('OItemColorName', merging=True,
        subtotal=['Sum1', 'Sum2', 'Sum3', 'Sum4', 'Sum5', 'Sum6', 'Sum7', 'Sum8', 'Sum9', 'Sum10', 'Sum11', 'Sum12', 'Sum13'],
        subtitle=my_second_header)

# печатаю таблицу в отчёт
cur_row = rep.apply_table(table, first_row=cur_row)

# открываю отчет в программе по умолчанию для .xls
rep.launch_excel()
