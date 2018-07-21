#!/usr/bin/python
# -*- coding: utf-8 -*-

from xlstable import *
import config
import pymssql

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