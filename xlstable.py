#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlsutils_apply import *
from openpyxl.utils import get_column_letter

from collections import namedtuple

class XLSTableColumnInfo:
    """Структура для хранения информации одного столбца данных таблицы
    """
    def __init__(self, fieldname, data_type = 'string', count=1, editable=False):
        self.fieldname  = fieldname
        self.count      = count
        self.type       = data_type
        self.editable   = editable

FieldStat = namedtuple('FieldStat', 'index format xls_start xls_end last_value last_value_row '\
                                    'hide_condition hide_flag')

class XLSTable:
    """Класс, инкапсулирующий информацию и методы отображения данных таблицы
    """
    def __init__(self, colinfo, data, row_height=30):
        self._columns = colinfo
        self._data = data

        self._row_height = row_height
        self._col_count = sum(coli.count for coli in colinfo)
        self._row_count = len(data)

        # используется в алгоритмах подитогов/подзаголовках/объединения строк с одинаковыми значениями
        self._fields = dict()

        index = 0
        xls_index = 0
        for coli in colinfo:
            self._fields[coli.fieldname] = FieldStat(index=index,
                                                     format=coli.type,
                                                     xls_start=xls_index,
                                                     xls_end=xls_index + coli.count - 1,
                                                     last_value=None,
                                                     last_value_row=None,
                                                     hide_condition=None,
                                                     hide_flag=None)
            index += 1
            xls_index += coli.count

        # используются в алгоритме объединения строк с одинаковыми значениями
        self._merged_hierarchy = []

    def add_hide_column_condition(self, fieldname, cond_func):
        self._fields[fieldname] = self._fields[fieldname]._replace(hide_condition=cond_func,
                                                                   hide_flag=True)

    def add_merge_column_hierarchy(self, field_hierarchy):
        self._merged_hierarchy = field_hierarchy

    def apply(self, ws, first_row, first_col):
        """Отображает непосредственно в XLS данные таблицы
        """
        def _save_last_values_and_merge(row, cur_row):
            """сохраняем предыдущие значения полей, объединяем ячейки с одинаковыми значениями
            """
            was_changed = False
            for fieldname in self._merged_hierarchy:
                col = self._fields[fieldname]
                new_value = row[col.index] if row else None

                if (new_value != col.last_value):
                    was_changed = True
                if was_changed and (col.last_value_row is not None):
                    apply_range(ws, col.last_value_row, first_col + col.xls_start,
                                    cur_row - 1,        first_col + col.xls_end, set_merge)
                if was_changed:
                    self._fields[fieldname] = col._replace(last_value=new_value,
                                                           last_value_row=cur_row)

        cur_row = first_row
        for data_row in self._data:
            ws.row_dimensions[cur_row].height = self._row_height

            for k, f in self._fields.items():
                if (f.xls_start - f.xls_end > 1):
                    apply_range(ws, cur_row, first_col + f.xls_start,
                                    cur_row, first_col + f.xls_end, set_merge)

                if (f.format not in ['int', 'currency', '3digit']) or (data_row[f.index] != 0):
                    ws.cell(row=cur_row, column=first_col + f.xls_start).value = data_row[f.index]

                if f.hide_condition is not None:
                    if not f.hide_condition(data_row[f.index]):
                        self._fields[k] = self._fields[k]._replace(hide_flag=False)

            _save_last_values_and_merge(data_row, cur_row)
            cur_row += 1

        _save_last_values_and_merge(None, cur_row)

        # скрываем все колонки, для которых не выполнились условия
        fields = [[fn.xls_start, fn.xls_end] for fn in self._fields.values() if fn.hide_flag]
        for fstart, fend in fields:
            for i in range(fstart, fend + 1):
                ws.column_dimensions[get_column_letter(first_col + i)].hidden = True

        # apply format and alignment
        cur_col = first_col
        for coli in self._columns:
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
