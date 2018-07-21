#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlsutils import *
from xlsutils_apply import *
from xlscolor import *
from openpyxl.utils import get_column_letter

class XLSTableHeaderColumn:
    """Структура для хранения информации одного столбца (ячейки) шапки таблицы
    """
    def __init__(self, title='', width=None, struct=[]):
        self.title = title
        self._struct = struct
        self.width = width if not struct else None
        self.count = self._struct_leaves_count()
        self.height = self._struct_height()
        self.titled_height = self._struct_titled_height()

    def _struct_leaves_count(self):
        """Подсчитывает количество колонок во всей структуре.
        Каждая вложенная структура уже посчитала свой count при инициализации.
        Внешние структуры инициализируются последними
        """
        return 1 if not self._struct else sum([thc.count for thc in self._struct])

    def _struct_height(self):
        """Подсчитывает глубину шапки
        """
        return 1 if not self._struct else 1 + max([thc.height for thc in self._struct])

    def _struct_titled_height(self):
        """Подсчитывает глубину видимой части шапки (колонки должны иметь title != '')
        """
        cnt = 1 if self.title != '' else 0
        if self._struct:
            return cnt + max([thc.titled_height for thc in self._struct if thc.title != ''], default=0)
        else: return cnt


class XLSTableHeader:
    """Класс, инкапсулирующий информацию и методы отображения шапки таблицы
    """
    def __init__(self, headers, bgcolor=Color.LT_GRAY.value):
        self._columns = headers
        self._bgcolor = bgcolor
        self._col_count = sum(hdr.count for hdr in headers)
        self._height = max([hdr.titled_height for hdr in headers], default = 0)

    def column_count(self):
        """Возвращает количество физических столбцов в таблице
        """
        return self._col_count

    def apply_widths(self, ws):
        """Применяет информацию о ширине столбцов из заголовка таблицы непосредственно к листу
        """
        pass

    def apply(self, ws, first_row, first_col):
        """Отображает непосредственно в XLS шапку таблицы
        """
        ws.row_dimensions[first_row].height = 50 if self._height == 1 else 24
        for i in range(1, self._height):
            ws.row_dimensions[first_row + i].height = 32

        cur_col = first_col
        for thc in self._columns:
            if thc.count > 1:
                ws.merge_cells(start_row=first_row, start_column=cur_col,
                               end_row=first_row,   end_column=cur_col + thc.count - 1)
            else: ws.merge_cells(start_row=first_row, start_column=cur_col,
                                 end_row=first_row + self._height - 1, end_column=cur_col)
            if not thc.width is None:
                ws.column_dimensions[get_column_letter(cur_col)].width = thc.width

            ws.cell(row=first_row, column=cur_col).value = thc.title
            cur_col += thc.count

        cr = get_xlrange(first_row,                    first_col,
                         first_row + self._height - 1, cur_col - 1)
        apply_xlrange(ws, cr, set_borders)
        apply_xlrange(ws, cr, set_outline, border_style='medium')
        apply_xlrange(ws, cr, set_font, bold=True)
        apply_xlrange(ws, cr, set_alignment)
        apply_xlrange(ws, cr, set_fill, color=self._bgcolor)

        return first_row + self._height
