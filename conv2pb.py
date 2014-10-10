# -*- coding: utf-8 -*-

import sys
import os
import getopt

abspath = os.path.abspath(sys.argv[0])
sys.path.append(os.path.dirname(abspath) + '/dep')
import xlrd

TMAP = {}
TMAP[1] = int
TMAP[2] = long
TMAP[3] = int
TMAP[4] = long
TMAP[5] = float
TMAP[6] = float
TMAP[7] = bool
# 8: enumerate
TMAP[9] = str
# 10: message

LMAP = {}
LMAP[1] = 'optional'
LMAP[2] = 'required'
LMAP[3] = 'repeated'

MOD = []

PROTOC = 'protoc'
PYTHON_PATH = './'

######################################################################

def load_enum(module, enum_key):
    if module.__dict__.has_key(enum_key):
        return eval('module.%s' % enum_key)
    else:
        for mod in MOD:
            if mod.__dict__.has_key(enum_key):
                return eval('mod.%s' % enum_key)
    print 'cant find enum define [%s]' % enum_key
    quit()

######################################################################

def get_field_name(prototype, key):
    if key not in prototype.fields_by_name:
        ds = key.split('.')
        if len(ds) == 1:
            return (None, None)
        if ds[0] not in prototype.fields_by_name:
            print '%s -> %s not found' % (key, ds[0])
            return (None, None)
        return (ds[0], '.'.join(ds[1:len(ds)]))
    return (key, None)

######################################################################

def load_field(module, field, obj, name, val):
    # 枚举类型
    if field.enum_type != None:
        val = load_enum(module, val)
    # 基本类型
    elif field.cpp_type in TMAP:
        visual_type = TMAP[field.cpp_type]
        val = visual_type(val)
    # 字符编码
    if type(val) == type(''):
        val = unicode(val, 'utf8')

    if LMAP[field.label] == 'repeated':
        getattr(obj, name).append(val)
    else:
        obj.__setattr__(name, val)

######################################################################

def load_stucture(module, field, obj, sub_field_name, val):
    prototype = field.message_type
    (name, sub_name) = get_field_name(prototype, sub_field_name)
    if name == None or name not in prototype.fields_by_name:
        print 'load %s fail' % sub_field_name
        quit()
    sub_field = prototype.fields_by_name[name]
    if sub_name == None:
        return load_field(module, sub_field, obj, name, val)
    else:
        return load_stucture(module, sub_field, getattr(obj, name), sub_name, val)

######################################################################

def load_data(module, row_field, row_obj, key, val, repeated_map):
    prototype = row_field.message_type

    # 是否是嵌套struct
    (name, sub_name) = get_field_name(prototype, key)
    if name == None:
        return
    item_field = prototype.fields_by_name[name]

    # 普通类型
    if item_field.message_type == None:
        return load_field(module, item_field, row_obj, name, val)

    # 嵌套structure
    else:
        if LMAP[item_field.label] == 'repeated':
            if name in repeated_map:
                item_field_obj = repeated_map[name]
            else:
                item_field_obj = getattr(row_obj, name).add()
                repeated_map[name] = item_field_obj
        else:
            item_field_obj = getattr(row_obj, name)
        return load_stucture(module, item_field, item_field_obj, sub_name, val)

######################################################################

def convert_sheet(sheet, module, output_path):
    config_name = sheet.name
    if not module.__dict__.has_key(config_name):
        print ('sheet[%s] config not found ignore' % config_name)
        return

    # protobuf data structure
    table = module.__dict__[config_name]()

    # 配置类型
    prototype = module.DESCRIPTOR.message_types_by_name[config_name]
    assert len(prototype.fields) > 0
    row_field = prototype.fields[0]

    # load table by rows
    for row in range(2, sheet.nrows):
        repeated_map = {}
        for col in range(sheet.ncols):
            # key 在第二行
            key = sheet.cell_value(1, col)
            val = sheet.cell_value(row, col)
            if type(val) == type(u''):
                val = val.encode('utf8')
            if val == '':
                continue
            if col == 0:
                row_objs = getattr(table, row_field.name)
                row_obj = row_objs.add()
            load_data(module, row_field, row_obj, key, val, repeated_map)

    # output bin
    output_file = '%s/%s.pbin' % (output_path, config_name)
    try:
        fd = open(output_file, 'wb+')
    except:
        print '%s open fail' % output_file
        quit()
    fd.write(table.SerializePartialToString())
    fd.close()

    # check
    res = check(table, output_file)
    if res == False:
        print 'convert %s fail' % output_file
        return False
    print 'convert %s success' % output_file

    # log
    log_file = '%s/%s.log' % (output_path, config_name)
    from google.protobuf import text_format as mod
    try:
        fd = open(log_file, 'w')
    except:
        print '%s open fail' % log_file
        quit()
    fd.write(mod.MessageToString(table))
    fd.close()
    return True
######################################################################

def load_import(import_path):
    sys.path.append(PYTHON_PATH)
    for files in os.listdir(import_path):
        suffix = os.path.splitext(files)[1]
        if suffix == '.proto':
            cmd = '%s -I%s --python_out=%s %s/%s' % (PROTOC, import_path, PYTHON_PATH, import_path, files)
            os.system(cmd)
    for files in os.listdir(import_path):
        proto = os.path.splitext(files)[0]
        suffix = os.path.splitext(files)[1]
        if suffix == '.proto':
            mod = __import__('%s_pb2' % proto)
            MOD.append(mod)

######################################################################

def convert(excel, proto, output_path, import_path):
    proto_file = os.path.basename(proto)
    proto_path = os.path.dirname(proto)
    if proto_path == '':
        proto_path = '.'
    proto_name = os.path.splitext(proto_file)[0]
    proto_bin_name = '%s.pb' % proto_name

    # generate python-protobuf-interface
    if import_path != '':
        cmd = '%s -I%s -I%s --python_out=%s %s ' % (PROTOC, proto_path, import_path, PYTHON_PATH, proto)
    else:
        cmd = '%s -I%s --python_out=%s %s ' % (PROTOC, proto_path, PYTHON_PATH, proto)
    os.system(cmd)

    # import protocol
    sys.path.append(PYTHON_PATH)
    module = __import__('%s_pb2' % proto_name)

    # open excel
    try:
        book = xlrd.open_workbook(excel)
    except:
        print 'open xls[%s] fail' % excel
        quit()

    # get sheet one by one
    for sheet in book.sheets():
        if convert_sheet(sheet, module, output_path) == False:
            print '%s convert fail\n' % excel
            quit()
    print '%s convert success\n' % excel

######################################################################

def check(message, file_name):
    try:
        fd = open(file_name, 'rb')
        message.ParseFromString(fd.read())
    except:
        print 'open check file[%s] fail' % file_name
        return False
    return True

######################################################################

def usage():
    print '''
    conv2pb.py usage:
        -f, --excel_file    <excel file name>               [required]
        -p, --proto         <protobuf decript protocol>     [required]
        -I, --import_path   <protobuf import path>          [optional]
        -O, --output_path   <convert result's path>         [optional]
    '''
    quit()

######################################################################

if __name__ == '__main__':
    sopt = 'f:p:I:O:'
    lopt = ['excel_file=', 'proto=', 'import_path=', 'output_path=']
    try:
        opts, args = getopt.getopt(sys.argv[1:], sopt, lopt)
    except:
        usage()

    excel = ''
    proto = ''
    import_path = ''
    output_path = '.'
    for opt, val in opts:
        if opt in ('-f', '--excel_file'):
            excel = val
        elif opt in ('-p', '--proto'):
            proto = val
        elif opt in ('-I', '--import_path'):
            import_path = val
        elif opt in ('-O', '--output_path'):
            output_path = val

    if excel == '' or proto == '':
        usage()

    sys.path.append('.')

    if import_path != '':
        load_import(import_path)

    convert(excel, proto, output_path, import_path)

