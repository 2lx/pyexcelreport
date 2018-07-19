#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
import sys
import subprocess
import tempfile

def temporary_file( filetemplate='sample' ):
    'Генерирует уникальное для временного каталога имя файла по переданному шаблону. Возвращает путь к файлу'
    cnt = 1
    while True:
        newfilename = '{0:s}{1:s}{2:s}{3:0>2d}.xlsx'.format( tempfile.gettempdir(), os.sep, filetemplate, cnt )
        if os.path.isfile(newfilename):
            cnt+=1
        else:
            break

    return newfilename

def open_file( filename ):
    'Opens file in app which is associated with file\'s extension'
    if sys.platform == 'win32':
        os.startfile( filename )
    else:
        opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
        subprocess.call( [opener, filename] )
