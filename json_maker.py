# -*- coding: utf-8 -*-

from base import *

JSON_TAB        = '  '
JSON_LINE_BREAK = '\n'
JSON_BLANK      = ''

JSON_DIRECTOR   = 'json'

class JSONMaker(Base):

    _output_json_texts = {}

    def __init__(self):
        Base.__init__(self)
        self._parse_to_json()

    def _table_to_json(self, sheet_name, output_name):
        if sheet_name == TEMPLATE_TABLE_NAME:
            return JSON_BLANK

        table = self._try_get_sheet_by_name(sheet_name)
        column_names = table.row_values(TABLE_COLUMN_NAME_ROW_INDEX)
        column_types = table.row_values(TABLE_COLUMN_TYPE_ROW_INDEX)

        text = JSON_BLANK
        for row_index in range(TABLE_OFFSET_ROW_NUM, table.nrows):
            row_text = JSON_BLANK
            for column_index in range(table.ncols):
                cell = table.cell(row_index, column_index)
                cell_value = cell.value
                cell_type = column_types[column_index]
                if cell.ctype == xlrd.book.XL_CELL_NUMBER:
                    if cell_type == COLUMN_INT_TYPE or cell_type == COLUMN_UINT_TYPE:
                        cell_value = int(cell_value)
                if len(row_text) > 0:
                    row_text += ','
                if cell_type == COLUMN_STRING_TYPE:
                    cell_value = str.format('"{}"', cell_value)
                row_text += str.format('{}', cell_value)
            row_text = str.format('[{}]', row_text)
            if len(text) > 0:
                text += ','
            text += row_text
        text = str.format('"{}":[{}]', output_name, text)
        return text

    def _parse_to_json(self):
        for row in self._output_config_table:
            output_file_name = row[CONFIG_OUTPUT_FILE_NAME_COLUMN_INDEX]
            output_json_text = self._output_json_texts.get(output_file_name)
            if output_json_text == None:
                output_json_text = JSON_BLANK
            if len(output_json_text) > 0:
                output_json_text += ','
            output_json_text += self._table_to_json(
                row[CONFIG_SHEET_NAME_COLUMN_INDEX], row[CONFIG_OUTPUT_NAME_COLUMN_INDEX])
            self._output_json_texts[output_file_name] = output_json_text

    def exportJSON(self):
        for key in self._output_json_texts.keys():
            value = self._output_json_texts.get(key)
            if value != None:
                value = '{' + value + '}'
                self._write(JSON_DIRECTOR, key, value)

if __name__ == '__main__':
    json = JSONMaker()
    json.exportJSON()