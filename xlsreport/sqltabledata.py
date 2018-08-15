#!/usr/bin/python
# -*- coding: utf-8 -*-

from .xlstable import *
import re
import datetime
import sys
from copy import copy

def maybe_sqlquoted(param):
    """Возвращает переданный параметр окруженный кавычками (если это необходимо в SQL)
    """
    UUID_PATTERN = re.compile(r'^[{]?([\da-f]{8}-([\da-f]{4}-){3}[\da-f]{12})[}]?$', re.IGNORECASE)

    matched = UUID_PATTERN.match(param)
    if matched:
        param = "'{{{0:s}}}'".format(matched.group(1))
    elif type(param) == [datetime.date]:
        param = "'{0:s}'".format(param.strftime('%d/%m/%Y %H:%M:%S.%f'))
    elif isinstance(param, str):
        param = "'{0:s}'".format(param)
    return param


import pymssql
from . import config

class MSSql():
    def __init__(self, server=config.mssql_server,
                       user=config.db_login,
                       password=config.db_password,
                       database=config.db_catalog):
        self.conn = pymssql.connect(server=server,
                                    user=user,
                                    password=password,
                                    database=database,
                                    autocommit=True)

    def __del__(self):
        self.conn.close

    def get_table_data(self, sqlquery, table_info):
        """Возвращает все поля из результатов запроса в формате списка значений (в порядке полей из table_info)
        """
        cursor = self.conn.cursor(as_dict=True)

        print("Выполняется запрос: '{0:s}'".format(sqlquery))
        cursor.execute(sqlquery)
        table_data = []

        for row in cursor:
            row_data = ()
            for ti in table_info:
                row_data += (row[ti.fname],) if (ti.fname != '') and (not ti.fname.startswith('__')) else ('',)
            table_data.append(row_data)

        return table_data

    def get_dict_data(self, sqlquery):
        """Возвращает все поля из результатов запроса в формате списка словарей
        """
        cursor = self.conn.cursor(as_dict=True)

        print("Выполняется запрос: '{0:s}'".format(sqlquery))
        cursor.execute(sqlquery)
        dict_data = []

        for row in cursor:
#            one_dict = dict()
#            for fld in row.keys():
#                one_dict[fld] = row[fld]
            dict_data.append(copy(row))

        return dict_data
