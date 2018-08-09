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

"""Самая простая функция - вывести строку в ячейку, либо в неск. объединенных ячеек в одной строке.
Это лейблы, реализация в xlslable.py
"""
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
Реализация в xlstableheader.py
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

"""Теперь можно печатать сам отчёт
Реализация в xlstable.py
Сначала нужно получить или задать данные для него. Для получения из MS SQL сервера есть методы из
модуля sqltabledata, но здесь не будут рассматриваться т.к. это вроде как отдельная либа да и не
переносимо. В общем случае данные задаются двумя списками - список полей типа TF (псевдоним) и
сам список данных - это список из кортежей с одной структурой, заданный первым списком.
В 1 структуре - порядок важен
1 - название поля, в SQL используется в запросе.
2 - форматирование прописано пока только для string, int, 3digit, currency
3 - сколько столбцов занимает поле в Excel (по умолчанию 1 можно не задавать)
Поле '' - в SQL вернет пустое значение
"""
# задаю структуру контекста отчёта
table_info = (\
        TF('ArticleGlobalCode',    'string', 1),
        TF('OItemColorName',       'string', 2),
        TF('Sum1',                 'int',    1),
        TF('Sum2',                 'int',    1),
        TF('Sum3',                 'int',    1),
        TF('Sum4',                 'int',    1),
        TF('',                     'string', 1),
        )

# указываю данные для контекста
table_data = [
    [ 'MSH05435', 'черный', 50, 0, 150, 100,  '1-50' ],
    [ 'MSH05435', 'черный', 50, 0, 150, 200, '31-50' ],
    [ 'MSH05436', 'белый',   0, 0, 150,   0,  'данные' ],
    [ 'MSH05437', 'белый',   0, 0, 150, 100,  '1-50' ],
    [ 'MSH05437', 'черный',  0, 0, 150, 100,  '1-50' ],
    [ 'MSH05437', 'черный',  0, 0, 150, 100,  '1-50' ],
    [ 'MSH05437', 'красный',50, 0, 150, 100,  'r0' ],
    [ 'MSH05437', 'желтый', 50, 0, 150, 200,  '1-50' ] ]

table = XLSTable(table_info, table_data, row_height=20)

# печатаю таблицу в отчёт
cur_row += 1
cur_row = rep.print_label(XLSLabel('Базовая таблица данных', LabelHeading.h3),
                          first_row=cur_row, col_count=max_col)
cur_row = rep.print_table(table, first_row=cur_row)

""" Таблица поддерживает объединение ячеек по вертикали. Задается это через иерархию полей, которая
по сути должна совпадать с сортировкой данных в таблице. Пример объединения ниже
"""

table2 = XLSTable(table_info, table_data, row_height=20)
table2.hierarchy_append('ArticleGlobalCode', merging=True)
table2.hierarchy_append('OItemColorName', merging=True)

cur_row += 1
cur_row = rep.print_label(XLSLabel('Объединяем одинаковые значения столбцов по вертикали', LabelHeading.h3),
                          first_row=cur_row, col_count=max_col)
cur_row = rep.print_table(table2, first_row=cur_row)

""" Можно скрыть столбцы, для которых выполнилось условие скрытия для каждого значения поля.
Пример - добавим условие для столбцов Sum1 - Sum4
"""

table3 = XLSTable(table_info, table_data, row_height=20)

fn = lambda x: x == 0
for i in range(1, 5):
    table3.add_hide_column_condition("Sum{0:d}".format(i), fn)

# не применяю, т.к. подействует на весь отчёт
#  cur_row += 1
#  cur_row = rep.print_label(XLSLabel('Скрою значение столбцов', LabelHeading.h3),
#                            first_row=cur_row, col_count=max_col)
#  cur_row = rep.print_table(table3, first_row=cur_row)


""" Определяю функции для отрисовки подзаголовков таблицы
"""
def my_header_func(ws, row_data, cur_row, first_col):
    # в функции подзаголовка можно рисовать как обычно, и в том числе вызывать методы report
    colheaders = [ THC("р. {0:d}".format(i)) for i in range(1, 5) ]
    my_subtitle_header = XLSTableHeader( columns=[
            THC('Составной подзаголовок модели', struct=colheaders)],
            row_height=16 )

    cur_row = rep.print_tableheader(my_subtitle_header, first_row=cur_row, first_col=first_col + 3)
    return cur_row

table4 = XLSTable(table_info, table_data, row_height=20)

# указываю поля, которые участвуют в группировках равных значений в порядке приоритета
table4.hierarchy_append('ArticleGlobalCode', subtitle=my_header_func)

cur_row += 1
cur_row = rep.print_label(XLSLabel('Таблица с подзаголовками', LabelHeading.h3),
                          first_row=cur_row, col_count=max_col)
cur_row = rep.print_table(table4, first_row=cur_row)

""" Подитоги таблицы
"""
table5 = XLSTable(table_info, table_data, row_height=20)

# указываю поля с подитогами, указываю поля с подзаголовками
table5.hierarchy_append('ArticleGlobalCode',
        subtotal=['Sum1', 'Sum2', 'Sum3', 'Sum4'])
table5.hierarchy_append('OItemColorName',
        subtotal=['Sum1', 'Sum2', 'Sum3', 'Sum4'])

cur_row += 1
cur_row = rep.print_label(XLSLabel('Таблица с подитогами', LabelHeading.h3),
                          first_row=cur_row, col_count=max_col)
cur_row = rep.print_table(table5, first_row=cur_row)

""" Теперь всё вместе
"""
table6 = XLSTable(table_info, table_data, row_height=20)

table6.hierarchy_append('ArticleGlobalCode', merging=True, subtitle=my_header_func,
        subtotal=['Sum1', 'Sum2', 'Sum3', 'Sum4'])
table6.hierarchy_append('OItemColorName',
        subtotal=['Sum1', 'Sum2', 'Sum3', 'Sum4'])

cur_row += 1
cur_row = rep.print_label(XLSLabel('Таблица всё вместе', LabelHeading.h3),
                          first_row=cur_row, col_count=max_col)
cur_row = rep.print_table(table6, first_row=cur_row)

""" На основании данных таблицы также можно рассчитать таблицу итогов. Нужно указать упорядоченный
список полей, по которым будет группировка и сортировка, а также список полей, по которым будут
рассчитываться суммы. Ниже пример итогов по первому столбцу
"""

table_total_info = (\
        TF('ArticleGlobalCode',    'string', 1),
        TF('Sum1',                 'int',    1),
        TF('Sum2',                 'int',    1),
        TF('Sum3',                 'int',    1),
        TF('Sum4',                 'int',    1),
        )

table_total_data = table.group_by_data(table_total_info,
        hierarchy=['ArticleGlobalCode'],
        sums=['Sum1', 'Sum2', 'Sum3', 'Sum4'])
table_total = XLSTable(table_total_info, table_total_data, row_height=24)

cur_row += 1
cur_row = rep.print_label(XLSLabel('Итоги по таблице', LabelHeading.h3),
                          first_row=cur_row, first_col=3, col_count=5)
cur_row = rep.print_table(table_total, first_row=cur_row, first_col = 3)

# открываю отчет в программе по умолчанию для .xls
rep.launch_excel()
