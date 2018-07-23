#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlsutils_apply import *
from openpyxl.utils import get_column_letter

class XLSTableColumnInfo:
    """Структура для хранения информации одного столбца данных таблицы
    """
    def __init__(self, fieldname, data_type = 'string', count=1, editable=False):
        self.fieldname  = fieldname
        self.count      = count
        self.type       = data_type
        self.editable   = editable

class XLSTable:
    """Класс, инкапсулирующий информацию и методы отображения данных таблицы
    """
    def __init__(self, colinfo, data, row_height=30):
        self._colinfo = colinfo
        self._data = data
        self._row_height = row_height

        self._col_count = sum(coli.count for coli in colinfo)
        self._row_count = len(data)

        # используются в алгоритме скрытия колонок по условию
        self._col_conditions = dict()
        self._col_hidden = dict()

        # используются в алгоритмах подитогов/подзаголовках/объединения строк с одинаковыми значениями
        self._fieldstat = dict()
        cur_index = 0
        for coli in colinfo:
            # информация в _fieldstat: номер столбца среди столбцов в данных, последнее значение, строка, с которой не менялись данные
            self._fieldstat[coli.fieldname] = [cur_index, None, None]
            cur_index += 1

        # используются в алгоритме объединения строк с одинаковыми значениями
        self._col_mergedhierarchy = []

    def add_hide_column_condition(self, field_name, cond_func):
        self._col_conditions[field_name] = cond_func
        self._col_hidden[field_name] = 1

    def add_merge_column_hierarchy(self, field_hierarchy):
        self._col_mergedhierarchy = field_hierarchy

    def apply(self, ws, first_row, first_col):
        """Отображает непосредственно в XLS данные таблицы
        """
        def _conditionally_hide_columns():
            """Скрывает колонки для которых для каждого значения в столбце выполнились условия скрытия
            """
            hidden_columns = [k for k, i in self._col_hidden.items() if i == 1]

            cur_col = first_col
            for coli in self._colinfo:
                if coli.fieldname in hidden_columns:
                    ws.column_dimensions[get_column_letter(cur_col)].hidden = True
                cur_col += coli.count

        def _save_last_values_and_merge(row, coli, col_index, cur_row, cur_col):
            # сохраняем предыдущее значение всех полей, которые проверяем в алгоритмах
            if coli.fieldname in self._fieldstat.keys():
                # если изначальное значение ещё не задано
                if  (self._fieldstat[coli.fieldname][1] is None):
                    self._fieldstat[coli.fieldname][1] = row[col_index]
                    self._fieldstat[coli.fieldname][2] = cur_row
                # возможно значение изменилось в самом столбце, либо в иерархии выше него
                elif (coli.fieldname in self._col_mergedhierarchy):
                    nothing_changes = True
                    for fh in self._col_mergedhierarchy:
                        if row[self._fieldstat[fh][0]] != self._fieldstat[fh][1]:
                            nothing_changes = False
                            break
                        if fh == coli.fieldname: break

                    # выяснили, что поле или поля от которых поле зависит изменились
                    if not nothing_changes:
                        merge_start_row = self._fieldstat[coli.fieldname][2]
                        print("({0:d},{1:d}) - ({2:d},{3:d})".format(merge_start_row, cur_col, cur_row, cur_col + coli.count - 1))
                        ws.merge_cells(start_row=merge_start_row, start_column=cur_col,
                                       end_row=cur_row - 1,       end_column=cur_col + coli.count - 1)

                        self._fieldstat[coli.fieldname][1] = row[col_index]
                        self._fieldstat[coli.fieldname][2] = cur_row

        cur_row = first_row
        for row in self._data:
            ws.row_dimensions[cur_row].height = self._row_height

            cur_col = first_col
            col_index = 0
            for coli in self._colinfo:
                if (coli.count > 1):
                    ws.merge_cells(start_row=cur_row, start_column=cur_col,
                                   end_row=cur_row,   end_column=cur_col + coli.count - 1)

                if (coli.type not in ['int', 'currency', '3digit']) or (row[col_index] != 0):
                    ws.cell(row=cur_row, column=cur_col).value = row[col_index]

                if coli.fieldname in self._col_conditions.keys():
                    if not self._col_conditions[coli.fieldname](row[col_index]):
                        self._col_hidden[coli.fieldname] = 0

                _save_last_values_and_merge(row, coli, col_index, cur_row, cur_col)

                # работа в цикл
                cur_col += coli.count
                col_index += 1
            cur_row += 1

        erow = [None] * len(self._colinfo)
        cur_col = first_col
        col_index = 0
        for coli in self._colinfo:
            _save_last_values_and_merge(erow, coli, col_index, cur_row, cur_col)
            cur_col += coli.count
            col_index += 1
        _conditionally_hide_columns()

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
