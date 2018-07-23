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

FieldStat = namedtuple('FieldStat', 'index xls_start xls_end last_value last_value_row')

class XLSTable:
    """Класс, инкапсулирующий информацию и методы отображения данных таблицы
    """
    def __init__(self, colinfo, data, row_height=30):
        self._columns = colinfo
        self._data = data

        self._row_height = row_height
        self._col_count = sum(coli.count for coli in colinfo)
        self._row_count = len(data)

        # используются в алгоритме скрытия колонок по условию
        self._col_conditions = dict()
        self._col_hidden = dict()

        # используется в алгоритмах подитогов/подзаголовках/объединения строк с одинаковыми значениями
        self._fields = dict()

        index = 0
        xls_index = 0
        for coli in colinfo:
            self._fields[coli.fieldname] = FieldStat(index=index,
                                                        xls_start=xls_index,
                                                        xls_end=xls_index + coli.count - 1,
                                                        last_value=None,
                                                        last_value_row=None)
            index += 1
            xls_index += coli.count

        # используются в алгоритме объединения строк с одинаковыми значениями
        self._merged_hierarchy = []

    def add_hide_column_condition(self, field_name, cond_func):
        self._col_conditions[field_name] = cond_func
        self._col_hidden[field_name] = 1

    def add_merge_column_hierarchy(self, field_hierarchy):
        self._merged_hierarchy = field_hierarchy

    def apply(self, ws, first_row, first_col):
        """Отображает непосредственно в XLS данные таблицы
        """
        def _conditionally_hide_columns():
            """Скрывает колонки для которых для каждого значения в столбце выполнились условия скрытия
            """
            hidden_columns = [k for k, i in self._col_hidden.items() if i == 1]

            cur_col = first_col
            for coli in self._columns:
                if coli.fieldname in hidden_columns:
                    ws.column_dimensions[get_column_letter(cur_col)].hidden = True
                cur_col += coli.count

        def _save_last_values_and_merge(row, cur_row):
            """сохраняем предыдущие значения всех полей, которые надо объединять, объединяем их
            """
            was_changed = False
            for fieldname in self._merged_hierarchy:
                col = self._fields[fieldname]
                new_value = row[col.index] if row else None

                if (new_value != col.last_value): was_changed = True

                if was_changed and (col.last_value_row is not None):
                    start_xlscol = first_col + col.xls_start
                    end_xlscol   = first_col + col.xls_end
                    ws.merge_cells(start_row=col.last_value_row, start_column=start_xlscol,
                                   end_row=cur_row - 1,          end_column=end_xlscol)
                   # print("({0:d},{1:d}) - ({2:d},{3:d})".format(col.last_value_row, xls_start,
                   #                                              cur_row - 1, xls_end))
                if was_changed:
                    self._fields[fieldname] = col._replace(last_value=new_value,
                                                              last_value_row=cur_row)

        cur_row = first_row
        for data_row in self._data:
            ws.row_dimensions[cur_row].height = self._row_height

            cur_col = first_col
            col_index = 0
            for coli in self._columns:
                if (coli.count > 1):
                    ws.merge_cells(start_row=cur_row, start_column=cur_col,
                                   end_row=cur_row,   end_column=cur_col + coli.count - 1)

                if (coli.type not in ['int', 'currency', '3digit']) or (data_row[col_index] != 0):
                    ws.cell(row=cur_row, column=cur_col).value = data_row[col_index]

                if coli.fieldname in self._col_conditions.keys():
                    if not self._col_conditions[coli.fieldname](data_row[col_index]):
                        self._col_hidden[coli.fieldname] = 0

                # работа в цикле
                cur_col += coli.count
                col_index += 1

            _save_last_values_and_merge(data_row, cur_row)
            cur_row += 1

        _save_last_values_and_merge(None, cur_row)

        _conditionally_hide_columns()

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
