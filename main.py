# -*- coding: utf-8 -*-
# 使用方法参照xmltool

import sys
import os

reload(sys)
sys.setdefaultencoding('utf8')

JSON_MAKER_FILE = 'json_maker.py'
CPP_MAKER_FILE = 'cpp_maker.py'

if __name__ == '__main__':
    xml_file_path = os.path.abspath(JSON_MAKER_FILE)
    if not os.path.exists(xml_file_path):
        print str.format('ERROR: not exists the "{}" file', JSON_MAKER_FILE)
        sys.exit()

    cpp_file_path = os.path.abspath(CPP_MAKER_FILE)
    if not os.path.exists(cpp_file_path):
        print str.format('ERROR: not exists the "{}" file', CPP_MAKER_FILE)
        sys.exit()

    if len(sys.argv) == 1:
        print 'ERROR: please specify the *.xlsx or *.xls file as the paramter'
    else:
        os.system(str.format('python {} {}', JSON_MAKER_FILE, sys.argv[1]))
        os.system(str.format('python {} {}', CPP_MAKER_FILE, sys.argv[1]))