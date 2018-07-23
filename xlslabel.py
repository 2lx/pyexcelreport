#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlsutils_apply import *
from enum import Enum

class LabelHeading(Enum):
    """Параметры отображения метки
          Высота строки, жирный текст, размер текста, гор. выр-е, верт выр-е
    """
    h1 = (30, True,   14, 'center', 'center')
    h2 = (24, True,   12, 'center', 'center')
    h3 = (20, True,   11, 'center', 'center')
    h4 = (20, False,  11, 'center', 'center')
    h5 = (18, False,  11, 'left',   'top')

class XLSLabel:
    """Класс, инкапсулирующий информацию и методы отображения заголовка
    """
    def __init__(self, title, heading=LabelHeading.h1):
        self.title=title
        self.heading=heading

    def apply(self, ws, first_row, first_col=1, col_count=1):
        ws.row_dimensions[first_row].height = self.heading.value[0]
        apply_range(ws, first_row, first_col, first_row, first_col + col_count -1, set_merge)

        ws.cell(row=first_row, column=first_col).value = self.title

        range = CellRange(min_row=first_row, min_col=first_col,
                          max_row=first_row, max_col=first_col + col_count - 1)

        apply_xlrange(ws, range, set_font,
                bold=self.heading.value[1], size=self.heading.value[2])
        apply_xlrange(ws, range, set_alignment,
                horizontal=self.heading.value[3], vertical=self.heading.value[4])
