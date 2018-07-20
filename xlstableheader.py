#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlsutils import *
from xlscolor import *

class XLSTableHeaderColumn:
    """Структура для хранения информации одного столбца (ячейки) шапки таблицы
    """
    def __init__(self, title, xlscolumns=1):
        self.title=title
        self.xlscolumns=xlscolumns

class XLSTableHeader:
    """Класс, инкапсулирующий информацию и методы отображения шапки таблицы
    """
    def __init__(self, headers, bgcolor=Color.LT_GRAY.value):
        self._data = headers
        self._bgcolor = bgcolor
        self._col_count = sum(hdr.xlscolumns for hdr in headers)

    def column_count(self):
        """Возвращает количество физических столбцов в таблице
        """
        return self._col_count

    def apply(self, ws, first_row, first_col):
        """Отображает непосредственно в XLS шапку таблицы
        """
        ws.row_dimensions[first_row].height = 50

        cur_col = first_col
        for hdr in self._data:
            if hdr.xlscolumns > 1:
                ws.merge_cells(start_row=first_row, start_column=cur_col, \
                               end_row=first_row,   end_column=cur_col + hdr.xlscolumns - 1)

            ws.cell(row=first_row, column=cur_col).value = hdr.title
            cur_col += hdr.xlscolumns

        apply_border(   ws, first_row, first_col, end_col=cur_col - 1)
        apply_outline(  ws, first_row, first_col, end_col=cur_col - 1, border_style='medium')
        apply_font(     ws, first_row, first_col, end_col=cur_col - 1, bold=True)
        apply_alignment(ws, first_row, first_col, end_col=cur_col - 1)
        apply_fill(     ws, first_row, first_col, end_col=cur_col - 1, color=self._bgcolor)

