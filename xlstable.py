#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlsutils_apply import *
from openpyxl.utils import get_column_letter
from xlscolor import *

from recordclass import recordclass

class XLSTableColumnInfo:
    """Структура для хранения информации одного столбца данных таблицы
    """
    def __init__(self, fieldname, format = 'string', col_count=1, editable=False):
        self.fname    = fieldname
        self.ccount   = col_count
        self.format   = format
        self.editable = editable

FieldStruct = recordclass('FieldStruct', 'findex xls_start xls_end format '
                                         'last_value last_value_row changed '
                                         'hide_condition hide_flag merging preheader subtotal')

class XLSTable:
    """Класс, инкапсулирующий информацию и методы отображения данных таблицы
    """
    def __init__(self, colinfo, data, row_height=30):
        self._fields = dict()
        findex = 0
        cindex = 0
        for ci in colinfo:
            self._fields[ci.fname] = FieldStruct(
                        findex=findex,
                        xls_start=cindex, xls_end=cindex + ci.ccount - 1,
                        format=ci.format,
                        last_value=None, last_value_row=None, changed=False,
                        hide_condition=None, hide_flag=None,
                        merging=False, preheader=None, subtotal=None)
            findex += 1
            cindex += ci.ccount

        self._data = data
        self._hierarchy = []
        self._row_height = row_height
        self._col_count = sum(ci.ccount for ci in colinfo)
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
            """ставим флаг changed если значение поля в структуре hierarchy поменяло свое значение
            """
            was_changed = False
            fielddict = {fn:self._fields[fn] for fn in self._hierarchy}.items()
            for fieldname, f in fielddict:
                if (row is None) or (row[f.findex] != f.last_value):
                    was_changed = True
                if was_changed:
                   self._fields[fieldname].changed = True

        def _after_line_processing(row, cur_row):
            """сохраняем инфо о последних значениях для всех полей в структуре _hierarchy
            """
            fielddict = {fn:self._fields[fn] for fn in self._hierarchy if self._fields[fn].changed}.items()
            for fieldname, f in fielddict:
                self._fields[fieldname].last_value = row[f.findex] if row else None
                self._fields[fieldname].last_value_row = cur_row
                self._fields[fieldname].changed = False

        def _merge_previous_rows(cur_row):
            """объединяем ячейки с одинаковыми значениями
            """
            fields = [self._fields[fn] for fn in self._hierarchy if self._fields[fn].merging and self._fields[fn].changed]
            for f in fields:
                if (f.last_value_row is not None) and (f.last_value_row != cur_row - 1):
                    apply_range(ws, f.last_value_row, first_col + f.xls_start,
                                    cur_row - 1,      first_col + f.xls_end, set_merge)
                    apply_range(ws, f.last_value_row, first_col + f.xls_start,
                                    cur_row - 1,      first_col + f.xls_start, set_borders)

        def _make_subtotals(cur_row):
            """делаем подитоги
            """
            stlines = 0
            fields = [self._fields[fn] for fn in reversed(self._hierarchy) if self._fields[fn].subtotal and self._fields[fn].changed]
            for fch in fields:
                if (fch.last_value_row is not None) and (fch.last_value_row != cur_row - 1):
                    ws.row_dimensions[cur_row + stlines].height = 18
                    ws.cell(row=cur_row + stlines, column=first_col + fch.xls_start).value = 'Подитоги'
                    apply_cell(ws, cur_row + stlines, first_col + fch.xls_start, set_alignment)
                    apply_cell(ws, cur_row + stlines, first_col + fch.xls_start, set_font, bold=True)
                    #  apply_cell(ws, cur_row + stlines, first_col + fch.xls_start, set_fill, color=Color.LT_GRAY.value)
                    #  apply_range(ws, cur_row + stlines, first_col, cur_row + stlines, first_col + self._col_count - 1,
                            #  set_fill, color=Color.LT_GRAY.value)
                    for st in fch.subtotal:
                        f = self._fields[st]
                        if (f.xls_start - f.xls_end > 1):
                            apply_range(ws, cur_row + stlines, first_col + f.xls_start,
                                            cur_row + stlines, first_col + f.xls_end, set_merge)
                        formulae = "=SUBTOTAL(9,{0:s}{1:d}:{0:s}{2:d})".format(
                                get_column_letter(first_col + f.xls_start),
                                fch.last_value_row, cur_row - 1)
                        #  print(formulae)
                        ws.cell(row=cur_row + stlines, column=first_col + f.xls_start).value = formulae
                        apply_cell(ws, cur_row + stlines, first_col + f.xls_start, set_alignment, horizontal='right')
                        apply_cell(ws, cur_row + stlines, first_col + f.xls_start, set_format, format=f.format)
                        apply_cell(ws, cur_row + stlines, first_col + f.xls_start, set_borders)
                        apply_cell(ws, cur_row + stlines, first_col + f.xls_start, set_fill, color=Color.LT_GRAY.value)
                    stlines += 1

            return cur_row + stlines

        cur_row = first_row
        for data_row in self._data:
            _before_line_processing(data_row)
            _merge_previous_rows(cur_row)
            cur_row = _make_subtotals(cur_row)

            ws.row_dimensions[cur_row].height = self._row_height

            for fieldname, f in self._fields.items():
                if (f.xls_start - f.xls_end > 1):
                    apply_range(ws, cur_row, first_col + f.xls_start,
                                    cur_row, first_col + f.xls_end, set_merge)

                # если печатаю числа, не выводить нулевые значения
                if (f.format not in ['int', 'currency', '3digit']) or (data_row[f.findex] != 0):
                    ws.cell(row=cur_row, column=first_col + f.xls_start).value = data_row[f.findex]

                # обновляем флаг hide_flag чтобы скрыть в конце неиспользуемые колонки
                if (f.hide_condition is not None) and (not f.hide_condition(data_row[f.findex])):
                    self._fields[fieldname].hide_flag = False

            _after_line_processing(data_row, cur_row)

            # apply format and alignment
            for f in self._fields.values():
                xlr = get_xlrange(cur_row, first_col + f.xls_start, cur_row, first_col + f.xls_end)
                if f.format in ['int', 'currency', '3digit']:
                    apply_xlrange(ws, xlr, set_alignment, horizontal='right')
                    apply_xlrange(ws, xlr, set_format, format=f.format)
                else:
                    apply_xlrange(ws, xlr, set_alignment)

            # apply borders, outline, font
            cr = get_xlrange(cur_row, first_col, cur_row, first_col + self._col_count - 1)
            apply_xlrange(ws, cr, set_borders)
            #  #  apply_xlrange(ws, cr, set_outline, border_style='medium')
            apply_xlrange(ws, cr, set_font)

            cur_row += 1

        _before_line_processing(None)
        _merge_previous_rows(cur_row)
        cur_row = _make_subtotals(cur_row)

        # скрываем все колонки, для которых не выполнились условия
        fields = [[fn.xls_start, fn.xls_end] for fn in self._fields.values() if fn.hide_flag]
        for fstart, fend in fields:
            for i in range(fstart, fend + 1):
                ws.column_dimensions[get_column_letter(first_col + i)].hidden = True

        # apply borders, outline, font
        cr = get_xlrange(first_row, first_col, cur_row - 1, first_col + self._col_count - 1)
        #  apply_xlrange(ws, cr, set_borders)
        apply_xlrange(ws, cr, set_outline, border_style='medium')
        #  apply_xlrange(ws, cr, set_font)

        return cur_row
