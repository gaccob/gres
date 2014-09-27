# -*- coding: utf-8 -*-

import xlrd
import sys
import os
import getopt

TMAP = {}
TMAP[1] = int
TMAP[2] = long
TMAP[3] = int
TMAP[4] = long
TMAP[5] = float
TMAP[6] = float
TMAP[7] = bool
TMAP[9] = str

LMAP = {}
LMAP[1] = 'optional'
LMAP[2] = 'required'
LMAP[3] = 'repeated'

MOD = []

PROTOC = 'protoc'
TMP_FILE = []

######################################################################

def load_enum(module, enum_key):
    if module.__dict__.has_key(enum_key):
        return eval('module.%s' % val)
    else:
        for mod in MOD:
            if mod.__dict__.has_key(enum_key):
                return eval('mod.%s' % enum_key)
    print 'cant find enum define [%s]' % enum_key
    quit()

######################################################################

def load_data(module, prototype, msg, key, val):
    # if not defined in protocol, ignore
    if key not in prototype.fields_by_name:
        return
    field = prototype.fields_by_name[key]

    # value or enumerate
    if field.enum_type != None:
        value = load_enum(module, val)
    else:
        if not field.cpp_type in TMAP:
            print "protobuf field[%s] invalid cpp type" % key
            quit()
        visual_type = TMAP[field.cpp_type]
        value = visual_type(val)

    # repeated or not
    if LMAP[field.label] == 'repeated':
        obj = getattr(msg, key)
        obj.append(value)
    else:
        if type(value) == type(''):
            value = unicode(value, 'utf8')
        msg.__setattr__(key, value)

######################################################################

def convert_sheet_print(file_name, content):
    try:
        fd = open(file_name, 'w')
    except:
        print '%s open fail' % file_name
        quit()
    fd.write(content)
    fd.close()

######################################################################

def convert_sheet_check(message, file_name):
    try:
        fd = open(file_name, 'rb')
        message.ParseFromString(fd.read())
    except:
        print 'open check file[%s] fail' % file_name
        quit()

######################################################################

def convert_sheet(sheet, module, output_path):
    config_name = sheet.name
    if not module.__dict__.has_key(config_name):
        print ('sheett [%s] not found in protobuf' % config_name)
        return

    # protobuf data structure
    table = module.__dict__[config_name]()

    # proto type
    prototype = module.DESCRIPTOR.message_types_by_name[config_name]
    assert len(prototype.fields) > 0
    row_name = prototype.fields[0].name
    data_prototype = prototype.fields_by_name[row_name].message_type

    # load table by rows
    for row in range(2, sheet.nrows):
        for col in range(sheet.ncols):
            key = sheet.cell_value(1, col)
            val = sheet.cell_value(row, col)
            if type(val) == type(u''):
                val = val.encode('utf8')
            if val == '':
                continue
            if col == 0:
                rows = getattr(table, row_name)
                msg = rows.add()
            load_data(module, data_prototype, msg, key, val)

    # output bin file
    bin_file = '%s/%s.bin' % (output_path, config_name)
    convert_sheet_print(bin_file, table.SerializePartialToString())
    convert_sheet_check(table, bin_file)
    print 'convert %s success\n' % bin_file

    # output log file
    log_file = '%s/%s.log' % (output_path, config_name)
    from google.protobuf import text_format as mod
    convert_sheet_print(log_file, mod.MessageToString(table))

######################################################################

def load_import(import_path, proto):

    imports = []
    for f in os.listdir(import_path):
        if f == proto:
            continue
        import_proto = os.path.splitext(f)[0]
        suffix = os.path.splitext(f)[1]
        if suffix == '.proto':
            cmd = '%s --error_format=msvs -I%s --python_out=. %s/%s' % (PROTOC, import_path, import_path, f)
            os.system(cmd)
            imports.append(import_proto)

    for p in imports:
        mod = __import__('%s_pb2' % p)
        TMP_FILE.append(p)
        MOD.append(mod)
        print 'mod[%s] loaded' % p

######################################################################

def convert(excel, proto, output_path, import_path):

    proto_file = os.path.basename(proto)
    proto_path = os.path.dirname(proto)
    proto_name =  os.path.splitext(proto_file)[0]
    proto_bin_name = '%s.pb' % proto_name

    import_cmd = ''
    if import_path != '':
        import_cmd = '%s -I%s' % (import_cmd, import_path)
    if proto_path != '':
        import_cmd = '%s -I%s' % (import_cmd, proto_path)

    # generate python-protobuf-interface
    cmd = '%s --error_format=msvs %s --python_out=. %s ' % (PROTOC, import_cmd, proto)
    os.system(cmd)

    # import protocol
    module = __import__('%s_pb2' % proto_name)
    TMP_FILE.append(proto_name)

    # open excel
    try:
        book = xlrd.open_workbook(excel)
    except:
        print 'open xls[%s] fail' % excel
        quit()

    # get sheet one by one
    for sheet in book.sheets():
        convert_sheet(sheet, module, output_path)

    # delete temporary file
    for f in TMP_FILE:
        os.system('rm ./%s_pb2.py*' % f)

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
        load_import(import_path, proto)

    convert(excel, proto, output_path, import_path)

