#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlstable import *
import config
import pymssql
import re
import datetime

UUID_PATTERN = re.compile(r'^[\da-f]{8}-([\da-f]{4}-){3}[\da-f]{12}$', re.IGNORECASE)

def maybe_sqlquoted(param):
    if (UUID_PATTERN.match(param)) or (type(param) == [datetime.date]) or isinstance(param, str):
        param = "'{0:s}'".format(param)
    return param


def get_table_data(sqlquery, table_info):
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