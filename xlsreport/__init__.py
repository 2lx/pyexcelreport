#!/usr/bin/python
# -*- coding: utf-8 -*-

from .xlsreport import *

if sys.platform.startswith('win'):
    try:
        from .sqltabledata import *
    except ImportError as e:
        pass # module pymssql doesn't exists
