#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlsutils import *
from xlsutils_apply import *
from xlscolor import *
from openpyxl.utils import get_column_letter

class XLSTableHeaderColumn:
    """Структура для хранения информации одного столбца (ячейки) шапки таблицы
    """
    def __init__(self, title, count=1, width=None, struct=[]):
        self.title = title
        self._struct = struct

        self.count = count
        self.width = width if count == 1 and len(struct) == 0 else None

class XLSTableHeader:
    """Класс, инкапсулирующий информацию и методы отображения шапки таблицы
    """
    def __init__(self, headers, bgcolor=Color.LT_GRAY.value):
        self._data = headers
        self._bgcolor = bgcolor
        self._col_count = sum(hdr.count for hdr in headers)

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
            if hdr.count > 1:
                ws.merge_cells(start_row=first_row, start_column=cur_col, \
                               end_row=first_row,   end_column=cur_col + hdr.count - 1)

            if not hdr.width is None:
                ws.column_dimensions[get_column_letter(cur_col)].width = hdr.width

            ws.cell(row=first_row, column=cur_col).value = hdr.title
            cur_col += hdr.count

        range = CellRange(min_row=first_row, min_col=first_col, \
                          max_row=first_row, max_col=cur_col - 1)
        apply_xlrange(ws, range, set_borders)
        apply_xlrange(ws, range, set_outline, border_style='medium')
        apply_xlrange(ws, range, set_font, bold=True)
        apply_xlrange(ws, range, set_alignment)
        apply_xlrange(ws, range, set_fill, color=self._bgcolor)

