#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlsutils_apply import *
from openpyxl.utils import get_column_letter

from recordclass import recordclass

class XLSTableColumnInfo:
    """Структура для хранения информации одного столбца данных таблицы
    """
    def __init__(self, fieldname, data_type = 'string', count=1, editable=False):
        self.fieldname  = fieldname
        self.count      = count
        self.type       = data_type
        self.editable   = editable

FieldStruct = recordclass('FieldStat', 'index xls_start xls_end format '
                                       'last_value last_value_row changed '
                                       'hide_condition hide_flag merging preheader subtotal')

class XLSTable:
    """Класс, инкапсулирующий информацию и методы отображения данных таблицы
    """
    def __init__(self, colinfo, data, row_height=30):
        # используется в алгоритмах подитогов/подзаголовках/объединения строк с одинаковыми значениями
        self._fields = dict()
        index = 0
        xls_index = 0

        for coli in colinfo:
            self._fields[coli.fieldname] = \
                    FieldStruct(index=index, xls_start=xls_index, xls_end=xls_index + coli.count - 1,
                                format=coli.type, last_value=None, last_value_row=None, changed=False,
                                hide_condition=None, hide_flag=None,
                                merging=False, preheader=None, subtotal=None)
            index += 1
            xls_index += coli.count

        self._data = data
        self._hierarchy = []
        self._row_height = row_height
        self._col_count = sum(coli.count for coli in colinfo)
        self._row_count = len(data)

    def add_hide_column_condition(self, fieldname, cond_func):
        self._fields[fieldname].hide_condition = cond_func
        self._fields[fieldname].hide_flag = True

    def add_hierarchy_field(self, fieldname, merging=False, preheader=None, subtotal=None):
        self._hierarchy.append(fieldname)
        self._fields[fieldname].merging = merging
        self._fields[fieldname].preheader = preheader
        self._fields[fieldname].subtotal = subtotal

    def apply(self, ws, first_row, first_col):
        """Отображает непосредственно в XLS данные таблицы
        """
        def _before_line_processing(row):
            was_changed = False
            fielddict = {fn:self._fields[fn] for fn in self._hierarchy}.items()

            for fieldname, f in fielddict:
                if (row is None) or (row[f.index] != f.last_value):
                    was_changed = True
                if was_changed:
                   self._fields[fieldname].changed = True

        def _after_line_processing(row, cur_row):
            fielddict = {fn:self._fields[fn] for fn in self._hierarchy if self._fields[fn].changed}.items()
            for fieldname, f in fielddict:
                self._fields[fieldname].last_value = row[f.index] if row else None
                self._fields[fieldname].last_value_row = cur_row
                self._fields[fieldname].changed = False

        def _merge_fields(cur_row):
            """объединяем ячейки с одинаковыми значениями
            """
            fields = [self._fields[fn] for fn in self._hierarchy if self._fields[fn].merging and self._fields[fn].changed]
            for f in fields:
                if (f.last_value_row is not None):
                    apply_range(ws, f.last_value_row, first_col + f.xls_start,
                                    cur_row - 1,      first_col + f.xls_end, set_merge)

        cur_row = first_row
        for data_row in self._data:
            ws.row_dimensions[cur_row].height = self._row_height
            _before_line_processing(data_row)

            for k, f in self._fields.items():
                if (f.xls_start - f.xls_end > 1):
                    apply_range(ws, cur_row, first_col + f.xls_start,
                                    cur_row, first_col + f.xls_end, set_merge)

                # если печатаю числа, не выводить нулевые значения
                if (f.format not in ['int', 'currency', '3digit']) or (data_row[f.index] != 0):
                    ws.cell(row=cur_row, column=first_col + f.xls_start).value = data_row[f.index]

                if (f.hide_condition is not None) and (not f.hide_condition(data_row[f.index])):
                        self._fields[k].hide_flag = False

            _merge_fields(cur_row)

            _after_line_processing(data_row, cur_row)
            cur_row += 1

        _before_line_processing(None)
        _merge_fields(cur_row)

        # скрываем все колонки, для которых не выполнились условия
        fields = [[fn.xls_start, fn.xls_end] for fn in self._fields.values() if fn.hide_flag]
        for fstart, fend in fields:
            for i in range(fstart, fend + 1):
                ws.column_dimensions[get_column_letter(first_col + i)].hidden = True

        # apply format and alignment
        for f in self._fields.values():
            xlr = get_xlrange(first_row, first_col + f.xls_start, cur_row - 1, first_col + f.xls_end)
            if f.format in ['int', 'currency', '3digit']:
                apply_xlrange(ws, xlr, set_alignment, horizontal='right')
                apply_xlrange(ws, xlr, set_format, format=f.format)
            else:
                apply_xlrange(ws, xlr, set_alignment)

        # apply borders, outline, font
        cr = get_xlrange(first_row, first_col, cur_row - 1, first_col + self._col_count - 1)
        apply_xlrange(ws, cr, set_borders)
        apply_xlrange(ws, cr, set_outline, border_style='medium')
        apply_xlrange(ws, cr, set_font)

        return cur_row
