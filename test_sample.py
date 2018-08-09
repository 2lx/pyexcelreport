#!/usr/bin/python
# -*- coding: utf-8 -*-

import locale
import sys

if sys.platform.startswith('win'):
    locale.setlocale(locale.LC_ALL, 'rus_rus')
else:
    locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')

from xlsreport import *

THC = XLSTableHeaderColumn
TF  = XLSTableField

"""Создаем объект - книгу Excel. В нем будет 1 лист с переданным названием (не более 32 символов и
не всех). Сразу задаются параметры печати портрет или ландшафт и кол-во страниц в ширину. Типы
параметров печати можно посмотреть в xlsreport
"""
rep = XLSReport('Акт передачи образцов', print_setup=PrintSetup.PortraitW1)

# Количество колонок в отчёте. Этот параметр рассчитывается автоматически при задании шапки отчёта,
# это будет показано ниже. Чтобы не путаться зададим пока руками
max_col = 8

# добавляю информацию о времени и пользователе, сгенерировавшем отчёт
# После отрисовки любого объекта, он возвращает номер след. пустой строки на листе - cur_row
cur_row = rep.print_preamble(max_col)
assert cur_row == 2

# Самая простая функция - вывести строку в ячейку, либо в неск. объединенных ячеек в одной строке.
# Выводим все виды форматирования от h1 до h5
cur_row = rep.print_label(XLSLabel('Заголовок h1, шириной 16 столбцов', LabelHeading.h1),
                          first_row=cur_row, col_count=max_col)
cur_row = rep.print_label(XLSLabel('Заголовок h2, шириной 4 столбца от 3-го столбца', LabelHeading.h2),
                          first_row=cur_row, first_col = 3, col_count=4)
cur_row = rep.print_label(XLSLabel('Заголовок h3, шириной 16 столбцов', LabelHeading.h3),
                          first_row=cur_row, col_count=max_col)
cur_row = rep.print_label(XLSLabel('Заголовок h4, шириной 16 столбцов', LabelHeading.h4),
                          first_row=cur_row, col_count=max_col)
cur_row = rep.print_label(XLSLabel('Заголовок h5, шириной 16 столбцов', LabelHeading.h5),
                          first_row=cur_row, col_count=max_col)

# вывожу сразу 2 лейбла в строку (просто не обновляю cur_row)
rep.print_label(XLSLabel('Левый лейбл', LabelHeading.h5),
                          first_row=cur_row, col_count=4)
cur_row = rep.print_label(XLSLabel('Правый', LabelHeading.h5),
                          first_row=cur_row, first_col=5, col_count=max_col - 5 + 1)

"""Структура шапки отчёта
Каждый столбец листа Excel имеет свою ширину. Удобней задавать ее вместе с заголовком столбца,
т.к. структуры шапки таблицы может измениться.
Очевидно, что если на странице несколько шапок, то имеет смысл задавать ширину через самую
широкую либо через первую, у остальных шапку можно опускать.
Есть возможность делать столбец шириной несколько столбцов Excel, поэтому widths это список,
а не единичное значение.
"""
# Пример шапки таблицы 1. Указан также параметр Высоты строки шапки - 60
tableheader1 = XLSTableHeader( columns=(
        THC( 'Заголовок 1 шириной 20 символов',  widths=[20] ),
        THC( 'Заголовок 2 из 2х столбцов шириной 30 и 15',  widths=[30, 15] ),
        THC( 'Заголовок 3 из 4х столбцов шириной по 15', widths=[15]*4 ),
        THC( 'Заголовок 4. Высота строки = 60',  widths=[20] ),
        ), row_height=60 )

# Получаем информацию о количестве Excel-колонок в отчете
# Чтобы не считать вручную колонки, по информации, заданной в шапке, можно кол-во посчитать
max_col = tableheader1.column_count

# Применим информацию из шапки таблицы о ширине столбцов листа Excel к листу
rep.apply_column_widths(tableheader1)

# печатаю шапку отчёта
cur_row += 1
cur_row = rep.print_label(XLSLabel('Пример простой шапки отчёта', LabelHeading.h3),
                          first_row=cur_row, col_count=max_col)
cur_row = rep.print_tableheader(tableheader1, first_row=cur_row)

""" Теперь сделаем еще одну шапку, но посложнее - она уже будет состоять из 2х уровней.
Каждый столбец - объект класса с псевдонимом THC, в котором возможно указать член-данные struct типа
списка объектов THC. Параметр widths стоит указывать только у нижнего этажа (ширина колонки-заголовка
равна сумме ширин подколонок)
"""
tableheader2 = XLSTableHeader( columns=(
        THC( 'Заголовок 1',  widths=[20] ),
        THC( 'Заголовок 2',  widths=[30] ),
        THC( 'Составной заголовок', struct=[ THC('Подзаголовок', widths=[ 7]) ]*5 ),
        THC( 'Заголовок 4',  widths=[20] ),
        ) )

# печатаю шапку отчёта
cur_row += 1
cur_row = rep.print_label(XLSLabel('Пример составной шапки отчёта', LabelHeading.h3),
                          first_row=cur_row, col_count=max_col)
cur_row = rep.print_tableheader(tableheader2, first_row=cur_row)

# А что будет при разной составной вложенности?
tableheader3 = XLSTableHeader( columns=(
        THC( '1', struct=[THC('1.1'), THC('1.2', struct=[THC('1.2.1'), THC('1.2.2')])]),
        THC( '2' ),
        THC( '3', struct=[ THC('3.1'), THC(3.2) ]),
        THC( '4', struct=[ THC('4.1', widths=[1, 1]) ] ), # если не применять apply_column_widths, важно только сколько чисел
        ) )

cur_row += 1
cur_row = rep.print_label(XLSLabel('Пример составной сложной шапки отчёта', LabelHeading.h3),
                          first_row=cur_row, col_count=max_col)
cur_row = rep.print_tableheader(tableheader3, first_row=cur_row)

"""
# задаю структуру контекста отчёта
table_info = (\
        TF('ArticleGlobalCode',    'string', 1),
        TF('OItemColorName',       'string', 1),
        TF('Sum1',                 'int',    1),
        TF('Sum2',                 'int',    1),
        TF('Sum3',                 'int',    1),
        TF('Sum4',                 'int',    1),
        TF('Sum5',                 'int',    1),
        TF('Sum6',                 'int',    1),
        TF('Sum7',                 'int',    1),
        TF('Sum8',                 'int',    1),
        TF('Sum9',                 'int',    1),
        TF('Sum10',                'int',    1),
        TF('Sum11',                'int',    1),
        TF('Sum12',                'int',    1),
        TF('Sum13',                'int',    1),
        TF('',                     'string', 1),
        )

# указываю данные для контекста
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
def my_header_func(ws, row_data, cur_row, first_col):
    # в функции подзаголовка можно рисовать как обычно, и в том числе вызывать методы report
    colheaders = [ THC("р. {0:d}".format(i)) for i in range(1, 14) ]
    my_subtitle_header = XLSTableHeader( columns=[
            THC('Составной подзаголовок модели', struct=colheaders)],
            row_height=16 )

    cur_row = rep.print_tableheader(my_subtitle_header, first_row=cur_row + 1, first_col=first_col + 2)
    return cur_row

def my_second_header(ws, row_data, cur_row, first_col):
    # просто вывожу надпись
    return rep.print_label(XLSLabel('Подзаголовок цветомодели', LabelHeading.h3),
                          first_row=cur_row, first_col=first_col + 2, col_count=13)

# указываю поля, которые участвуют в группировках равных значений,
# указываю поля с подитогами, указываю поля с подзаголовками
# всё это указываю в порядке приоритета
table.hierarchy_append('ArticleGlobalCode', merging=True,
        subtotal=['Sum1', 'Sum2', 'Sum3', 'Sum4', 'Sum5', 'Sum6', 'Sum7', 'Sum8', 'Sum9', 'Sum10', 'Sum11', 'Sum12', 'Sum13'],
        subtitle=my_header_func)
table.hierarchy_append('OItemColorName', merging=True,
        subtotal=['Sum1', 'Sum2', 'Sum3', 'Sum4', 'Sum5', 'Sum6', 'Sum7', 'Sum8', 'Sum9', 'Sum10', 'Sum11', 'Sum12', 'Sum13'],
        subtitle=my_second_header)

# печатаю таблицу в отчёт
cur_row = rep.print_table(table, first_row=cur_row)
"""
# открываю отчет в программе по умолчанию для .xls
rep.launch_excel()
