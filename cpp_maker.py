# -*- coding: utf-8 -*-
# 该工具用于生成可读取用xml_generate.py生成的xml文件的cpp文件
# 被转换的excel文件需要遵守template.xlsx的标准
# 命令格式：python cpp_generate.py excel.xlsx
#         python cpp_generate.py ../directory/excel.xls
#         python cpp_generate.py directory/excel.xls

from base import *

TEMPLATE_MANAGER_H     = 'JSONManager.h'
TEMPLATE_MANAGER_CPP   = 'JSONManager.cpp'
TEMPLATE_LOADER        = 'JSONLoader.h'
TEMPLATE_DATA_H        = 'JSONData.h'
TEMPLATE_DATA_CPP      = 'JSONData.cpp'
TEMPLATE_DATA_ITEM_H   = 'JSONDataItem.h'
TEMPLATE_DATA_ITEM_CPP = 'JSONDataItem.cpp'

TEMPLATE_DIRECTORY = 'cpp_template'
WIRTE_DIRECTORY    = 'cpp'

CPP_INT_TYPE    = 'int'
CPP_UINT_TYPE   = 'unsigned int'
CPP_FLOAT_TYPE  = 'float'
CPP_DOUBLE_TYPE = 'double'
CPP_BOOL_TYPE   = 'bool'
CPP_STRING_TYPE = 'std::string'
CPP_LEFT_BRACE  = '{'
CPP_RIGHT_BRACE = '}'

HEADER = 'header'
CPP    = 'cpp'

ITEM_CLASS_NAME_SUFFIX = 'Item'

CPP_INT_LOAD_MARCO    = '{} = row[{}].asInt();'
CPP_UINT_LOAD_MARCO   = '{} = row[{}].asUInt();'
CPP_BOOL_LOAD_MARCO   = '{} = row[{}].asBool();'
CPP_FLOAT_LOAD_MARCO  = '{} = row[{}].asDouble();'
CPP_DOUBLE_LOAD_MARCO = '{} = row[{}].asDouble();'
CPP_STRING_LOAD_MARCO = '{} = row[{}].asString();'
CPP_DECLARE_VAR_MARCO = 'JSON_DECLARE_VAR({});'
CPP_LOAD_FILE_MARCO   = 'LOAD_JSON_FILE("{}");'
CPP_LOAD_DATA_MARCO   = 'LOAD_TABLE_DATA({});'

CPP_BRACE_FORMAT = '\t\t{0}\n\t\t{2}\n\t\t{1}'

COMMENT_POSITION = 30

class CPPMaker(Base):

    _cpp_templates = {}

    def __init__(self):
        Base.__init__(self)

    def _get_cpp_type(self, json_type):
        cpp_type = CPP_INT_TYPE
        if json_type == COLUMN_STRING_TYPE:
            cpp_type = CPP_STRING_TYPE
        elif json_type == COLUMN_FLOAT_TYPE:
            cpp_type = CPP_FLOAT_TYPE
        elif json_type == COLUMN_DOUBLE_TYPE:
            cpp_type = CPP_DOUBLE_TYPE
        elif json_type == COLUMN_UINT_TYPE:
            cpp_type = CPP_UINT_TYPE
        elif json_type == COLUMN_BOOL_TYPE:
            cpp_type = CPP_BOOL_TYPE
        return cpp_type

    def _get_marco(self, json_type):
        marco = CPP_INT_LOAD_MARCO
        if json_type == COLUMN_STRING_TYPE:
            marco = CPP_STRING_LOAD_MARCO
        elif json_type == COLUMN_FLOAT_TYPE:
            marco = CPP_FLOAT_LOAD_MARCO
        elif json_type == COLUMN_DOUBLE_TYPE:
            marco = CPP_DOUBLE_LOAD_MARCO
        elif json_type == COLUMN_UINT_TYPE:
            marco = CPP_UINT_LOAD_MARCO
        elif json_type == COLUMN_BOOL_TYPE:
            marco = CPP_BOOL_LOAD_MARCO
        return marco

    def _comment_space(self, clause):
        num = COMMENT_POSITION - len(clause)
        space = ''
        for index in range(num):
            space += ' '
        return space

    def _item_class_name(self, class_name):
        return str.format('{}{}', class_name, ITEM_CLASS_NAME_SUFFIX)

    def _data_item_var_list(self, column_types, column_names, column_titles):
        text = ''
        for index in range(len(column_types)):
            clause = str.format('{}{}{};', self._get_cpp_type(column_types[index]), ' ', column_names[index])
            text += str.format('\n\t{}{}// {}', clause, self._comment_space(clause), column_titles[index])
        return text

    def _column_index(self, index):
        if index == 0:
            index = '0u'
        return index

    def _data_item_load_list(self, column_types, column_names):
        text = ''
        for index in range(len(column_types)):
            load_list = str.format(self._get_marco(column_types[index]), 
                column_names[index], self._column_index(index))
            text += str.format('\n\t{}', load_list)
        return text

    def _data_item_h(self, class_name, column_types, column_names, column_titles):
        text = self._get_cpp_template(TEMPLATE_DATA_ITEM_H)
        var_list = self._data_item_var_list(column_types, column_names, column_titles)
        text = str.format(text, CPP_LEFT_BRACE, CPP_RIGHT_BRACE, class_name, var_list)
        return text

    def _data_item_cpp(self, class_name, column_types, column_names):
        template = self._get_cpp_template(TEMPLATE_DATA_ITEM_CPP)
        load_list = self._data_item_load_list(column_types, column_names)
        text = str.format(template, CPP_LEFT_BRACE, CPP_RIGHT_BRACE, class_name, load_list)
        return text

    def _data_item(self, sheet_name, output_name):
        if sheet_name == TEMPLATE_TABLE_NAME:
            return ''

        table = self._try_get_sheet_by_name(sheet_name)
        column_names = table.row_values(TABLE_COLUMN_NAME_ROW_INDEX)
        column_types = table.row_values(TABLE_COLUMN_TYPE_ROW_INDEX)
        column_titles = table.row_values(TABLE_COLUMN_TITLE_ROW_INDEX)

        ret = {}
        ret[HEADER] = self._data_item_h(output_name, column_types, column_names, column_titles)
        ret[CPP] = self._data_item_cpp(output_name, column_types, column_names)
        return ret

    def _data_item_list(self):
        ret = {}
        ret[HEADER] = ''
        ret[CPP] = ''
        for row in self._output_config_table:
            item = self._data_item(row[CONFIG_SHEET_NAME_COLUMN_INDEX], 
                self._item_class_name(row[CONFIG_OUTPUT_NAME_COLUMN_INDEX]))
            ret[HEADER] += str.format('{}\n\n', item[HEADER])
            ret[CPP] += str.format('{}\n', item[CPP])
        return ret

    def _export_data(self):
        header = self._get_cpp_template(TEMPLATE_DATA_H)
        cpp = self._get_cpp_template(TEMPLATE_DATA_CPP)

        item_list = self._data_item_list()
        header = str.format(header, item_list[HEADER])
        cpp = str.format(cpp, item_list[CPP])
        self._write_cpp(TEMPLATE_DATA_H, header)
        self._write_cpp(TEMPLATE_DATA_CPP, cpp)

    def _manager_cpp(self):
        texts = {}
        for row in self._output_config_table:
            file_name = row[CONFIG_OUTPUT_FILE_NAME_COLUMN_INDEX]
            text = texts.get(file_name)
            if text == None:
                marc = str.format(CPP_LOAD_FILE_MARCO, file_name)
                text = str.format('\t{}', marc)
            marc = str.format(CPP_LOAD_DATA_MARCO, row[CONFIG_OUTPUT_NAME_COLUMN_INDEX])
            text += str.format('\n\t\t\t{}', marc)
            texts[file_name] = text

        ret = ''
        for key in texts.keys():
            text = str.format(CPP_BRACE_FORMAT, CPP_LEFT_BRACE, CPP_RIGHT_BRACE, texts[key])
            ret += str.format('\n{}\n', text)
        return ret

    def _xml_manager_h(self):
        text = ''
        for row in self._output_config_table:
            marco = str.format(CPP_DECLARE_VAR_MARCO, row[CONFIG_OUTPUT_NAME_COLUMN_INDEX])
            text += str.format('\t{}\n', marco)
        return text

    def _export_manager(self):
        header = self._get_cpp_template(TEMPLATE_MANAGER_H)
        cpp = self._get_cpp_template(TEMPLATE_MANAGER_CPP)

        header = str.format(header, CPP_LEFT_BRACE, CPP_RIGHT_BRACE, 
            self._xml_manager_h(), ITEM_CLASS_NAME_SUFFIX)
        cpp = str.format(cpp, CPP_LEFT_BRACE, CPP_RIGHT_BRACE, self._manager_cpp())
        self._write_cpp(TEMPLATE_MANAGER_H, header)
        self._write_cpp(TEMPLATE_MANAGER_CPP, cpp)

    def _export_loader(self):
        loader = header = self._get_cpp_template(TEMPLATE_LOADER)
        loader = str.format(loader, CPP_LEFT_BRACE, CPP_RIGHT_BRACE)
        self._write_cpp(TEMPLATE_LOADER, loader)

    def _get_cpp_template(self, file_name):
        template = self._cpp_templates.get(file_name)
        if template == None:
            template = self._try_open_cpp_templaet_file(file_name)
            self._cpp_templates[file_name] = template
        return template

    def _cpp_template_directory_exists(self):
        return self._directory_exists(os.path.abspath(TEMPLATE_DIRECTORY))

    def _write_cpp(self, file_name, text):
        self._write(WIRTE_DIRECTORY, file_name, text)

    def _try_open_cpp_templaet_file(self, file_name):
        if self._cpp_template_directory_exists():
            file_path = os.path.abspath(os.path.join(TEMPLATE_DIRECTORY, file_name))
            if os.path.exists(file_path) and os.path.isfile(file_path):
                f = file(file_path, 'r')
                value = f.read()
                f.close()
                return value
            else:
                self._error(str.format('"{}"文件不存在', file_name))
                sys.exit()
        else:
            self._error(str.format('"{}"文件夹不存在', TEMPLATE_DIRECTORY))
            sys.exit()

    def exportCppCode(self):
        self._export_data()
        self._export_loader()
        self._export_manager()

if __name__ == '__main__':
    cpp = CPPMaker()
    cpp.exportCppCode()