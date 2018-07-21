#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlstable import *
import re
import datetime

import sys


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


if sys.platform.startswith('win'):
    import config
    import pymssql
    def get_mssql_data(sqlquery, table_info):
        conn = pymssql.connect( server=config.mssql_server,
                                user=config.db_login,
                                password=config.db_password,
                                database=config.db_catalog,
                                autocommit=True)
        cursor = conn.cursor(as_dict=True)

        cursor.execute(sqlquery)
        table_data = []

        for row in cursor:
            row_data = ()
            for ti in table_info:
                row_data += (row[ti.fieldname],) if ti.fieldname != '' else ('',)
            table_data.append(row_data)

        conn.close
        return table_data

