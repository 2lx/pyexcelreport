#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlsutils_apply import *

class XLSLabel:
    """Класс, инкапсулирующий информацию и методы отображения заголовка
    """
    def __init__(self, title, importance=1):
        self.title=title
        self.importance=importance

    def apply(self, ws, first_row, first_col=1, col_count=1):
        ws.row_dimensions[first_row].height = 30
        ws.merge_cells(start_row=first_row, start_column=first_col, \
                       end_row=first_row,   end_column=first_col + col_count - 1)

        ws.cell(row=first_row, column=first_col).value = self.title

        range = CellRange(min_row=first_row, min_col=first_col, \
                          max_row=first_row, max_col=first_col + col_count - 1)
        if self.importance == 1:
            apply_xlrange(ws, range, set_font, bold=True, size=14)
        if self.importance == 2:
            apply_xlrange(ws, range, set_font, bold=True, size=12)
        if self.importance == 3:
            apply_xlrange(ws, range, set_font, bold=True, size=11)
        if self.importance == 4:
            apply_xlrange(ws, range, set_font, bold=False, size=11)
        apply_xlrange(ws, range, set_alignment)

