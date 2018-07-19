#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlsfunc import *

class XLSTable:
    """Класс, автоматизирующий создание таблицы на листе книги XLS
    """

    def __init__(self, headers):
        self.data = headers

    def column_count(self):
        """Возвращает количество физических столбцов в таблице
        """
        return sum(dataopt[1] if len(dataopt) > 1 else 1 for dataopt in self.data)

    def apply_header(self, ws, first_row, first_col):
        """Отображает шапку таблицы
        """
        ws.row_dimensions[ first_row ].height = 50

        cur_col = first_col
        for titleopt in self.data:
            title = titleopt[0]

            cnt = 1
            if len(titleopt) > 1:
                cnt = titleopt[1]
                ws.merge_cells( start_row=first_row, start_column=cur_col, \
                                    end_row=first_row, end_column=cur_col + cnt - 1 )

            ws.cell( row=first_row, column=cur_col, value=title )
            cur_col+=cnt

        apply_border( ws, first_row, first_col, first_row, cur_col - 1 )
        apply_outline( ws, first_row, first_col, first_row, cur_col - 1, 'medium' )

