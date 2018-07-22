#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlsutils_apply import *
from openpyxl.utils import get_column_letter

class XLSTableColumnInfo:
    """Структура для хранения информации одного столбца данных таблицы
    """
    def __init__(self, fieldname, type = 'string', count=1, editable=False):
        self.fieldname  = fieldname
        self.count      = count
        self.type       = type
        self.editable   = editable

class XLSTable:
    """Класс, инкапсулирующий информацию и методы отображения данных таблицы
    """
    def __init__(self, colinfo, data):
        self._colinfo = colinfo
        self._data = data
        self._col_count = sum(coli.count for coli in colinfo)
        self._row_count = len(data)
        self._col_conditions = dict()
        self._col_hidden = dict()

    def column_count(self):
        """Возвращает количество физических столбцов в таблице
        """
        return self._col_count

    def add_hide_column_condition(self, field_name, cond_func):
        self._col_conditions[field_name] = cond_func
        self._col_hidden[field_name] = 1

    def apply(self, ws, first_row, first_col):
        """Отображает непосредственно в XLS данные таблицы
        """
        #print data
        cur_row = first_row
        for row in self._data:
            ws.row_dimensions[cur_row].height = 30

            cur_col = first_col
            col_index = 0
            for coli in self._colinfo:
                if coli.count > 1:
                    ws.merge_cells(start_row=cur_row, start_column=cur_col,
                                   end_row=cur_row,   end_column=cur_col + coli.columns - 1)

                if (coli.type not in ['int', 'currency', '3digit']) or (row[col_index] != 0):
                    ws.cell(row=cur_row, column=cur_col).value = row[col_index]

                if coli.fieldname in self._col_conditions.keys():
                    if not self._col_conditions[coli.fieldname](row[col_index]):
                        self._col_hidden[coli.fieldname] = 0

                cur_col += coli.count
                col_index += 1
            cur_row += 1

        # conditionally hide columns
        hidden_columns = [k for k, i in self._col_hidden.items() if i == 1]

        cur_col = first_col
        for coli in self._colinfo:
            if coli.fieldname in hidden_columns:
                print('hide')
                ws.column_dimensions[get_column_letter(cur_col)].hidden = True
            cur_col += coli.count

        # apply format and alignment
        cur_col = first_col
        for coli in self._colinfo:
            if coli.type in ['int', 'currency', '3digit']:
                apply_range(ws, first_row, cur_col, cur_row -1, cur_col,
                        set_alignment, horizontal='right')
                apply_range(ws, first_row, cur_col, cur_row -1, cur_col,
                        set_format, format=coli.type)
            else:
                apply_range(ws, first_row, cur_col, cur_row -1, cur_col, set_alignment)
            cur_col += coli.count

        # apply borders, outline, font
        cr = get_xlrange(first_row, first_col, cur_row - 1, cur_col - 1)
        apply_xlrange(ws, cr, set_borders)
        apply_xlrange(ws, cr, set_outline, border_style='medium')
        apply_xlrange(ws, cr, set_font)

        return cur_row
