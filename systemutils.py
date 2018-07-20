#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
import sys
import tempfile

def is_exists_and_locked(filepath):
    """Checks if a file is locked by opening it in append mode.
    If no exception thrown, then the file is not locked.
    """
    if not os.path.exists(filepath):
        return False

    file_object = None
    try:
        file_object = open(filepath, 'a')
        if file_object:
            file_object.close()
            return False

    except BaseException:
        return True

def temporary_file(filename_pattern='sample'):
    """Generates the unique name of file for given template
    in system user termporary directory'
    """
    for cnt in range(1, 100):
        newfilename = '{0:s}{1:s}{2:s}{3:0>2d}.xlsx'.format(\
                tempfile.gettempdir(), os.sep, filename_pattern, cnt)
        if not is_exists_and_locked(newfilename):
            break

    return newfilename

def open_file(filename):
    """Opens file in app which is associated with file's extension
    """
    if sys.platform == 'win32':
        os.startfile(filename)
    else:
        opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
        os.system("{0:s} {1:s} &".format(opener, filename))

