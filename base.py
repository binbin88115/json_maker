# -*- coding: utf-8 -*-

import sys
import os
import xlrd

reload(sys)
sys.setdefaultencoding('utf8')

CONFIG_TABLE_NAME   = '_config_'
TEMPLATE_TABLE_NAME = '_table_template_'

COLUMN_STRING_TYPE  = 'STRING'
COLUMN_INT_TYPE     = 'INT'
COLUMN_FLOAT_TYPE   = 'FLOAT'

CONFIG_TABLE_OFFSET_ROW_NUM          = 2
CONFIG_TABLE_OUTPUT                  = 'YES'
CONFIG_SHEET_NAME_COLUMN_INDEX       = 0
CONFIG_OUTPUT_NAME_COLUMN_INDEX      = 1
CONFIG_OUTPUT_FILE_NAME_COLUMN_INDEX = 2
CONFIG_ENABLE_OUTPUT_COLUMN_INDEX    = 3

ROW_ELEMENT_NAME = 'item'

TABLE_COLUMN_TYPE_ROW_INDEX = 2
TABLE_COLUMN_NAME_ROW_INDEX = 3
TABLE_OFFSET_ROW_NUM        = 4

XLS_FILE_PATH_ARGV_INDEX = 1

class Base:
    
    _output_config_table = []
    _ouput_directory = None
    _xls_data = None

    def __init__(self):
        self._try_open_xls_file(self._try_get_xls_file_path())
        self._parse_config_table()

    def _parse_config_table(self):
        config_table = self._try_get_sheet_by_name(CONFIG_TABLE_NAME)
        for row_index in range(CONFIG_TABLE_OFFSET_ROW_NUM, config_table.nrows):
            row = config_table.row_values(row_index)
            if row[CONFIG_ENABLE_OUTPUT_COLUMN_INDEX] == CONFIG_TABLE_OUTPUT:
                self._output_config_table.append(row)

    def _try_open_xls_file(self, xls_path):
        try:
            self._xls_data = xlrd.open_workbook(xls_path)
        except Exception:
            self._error('目标文件错误，请确定所指定的为excel文件')
            sys.exit()

    def _try_get_sheet_by_name(self, sheet_name):
        try:
            table = self._xls_data.sheet_by_name(sheet_name)
            return table
        except Exception:
            self._error(str.format('不存在名称为"{}"的表', sheet_name))
            sys.exit()

    def _try_get_path_from_argv(self):
        try:
            return sys.argv[XLS_FILE_PATH_ARGV_INDEX]
        except Exception:
            self._error('没有指定所要解析的excel文件')
            sys.exit()

    def _try_get_xls_file_path(self):
        xls_file_path = self._try_get_path_from_argv()
        xls_file_path = os.path.abspath(xls_file_path)
        if os.path.exists(xls_file_path) and os.path.isfile(xls_file_path):
            self._ouput_directory = os.path.dirname(xls_file_path)
            return xls_file_path
        else:
            self._error('所指定的excel文件的路径不正确')
            sys.exit()

    def _error(self, error_msg):
        self._print_utf8(str.format('ERROR：{}', error_msg))

    def _print_utf8(self, content):
        print content.decode('utf-8')

    def _directory_exists(self, path):
        if os.path.exists(path):
            return True
        return False

    def _write(self, directory, file_name, text):
        directory_path = os.path.abspath(os.path.join(self._ouput_directory, directory))
        if not self._directory_exists(directory_path):
            os.mkdir(directory_path)
        path = os.path.join(directory_path, file_name)
        f = file(path, 'w')
        f.write(text)
        f.close()
