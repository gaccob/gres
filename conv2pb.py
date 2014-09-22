# -*- coding: utf-8 -*-

import xlrd
import sys
import os

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

######################################################################

def load_data(module, prototype, msg, key, val):
    field = prototype.fields_by_name[key]

    # value or enumerate
    if field.enum_type != None:
        value = eval('module.%s' % val)
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
        msg.__setattr__(key, value)

######################################################################

def convert_sheet(sheet, module, output):
    config_name = sheet.name
    if not module.__dict__.has_key(config_name):
        print ('protobuf define[%s] not found' % config_name)
        quit()

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

    # output bin
    output_file = '%s/%s.bin' % (output, config_name)
    try:
        fd = open(output_file, 'wb+')
    except Except, e:
        print '%s open fail: %s' % (output_file, str(e))
        quit()
    fd.write(table.SerializePartialToString())

    # check
    try:
        check(table, output_file)
    except Exception, e:
        print '%s check fail: %s' % (output_file, str(e))
    print 'convert %s success' % output_file

######################################################################

def convert(excel, proto, output):

    # generate python-protobuf-interface
    os.system('protoc --error_format=msvs --python_out=%s %s' % (output, proto))

    # import protocol
    proto_file = os.path.basename(proto)
    proto_name =  os.path.splitext(proto_file)[0]
    sys.path.append(output)
    module = __import__('%s_pb2' % proto_name)

    # open excel
    try:
        book = xlrd.open_workbook(excel)
    except Exception, e:
        print str(e)
        quit()

    # get sheet one by one
    for sheet in book.sheets():
        convert_sheet(sheet, module, output)

######################################################################

def check(message, file_name):
    try:
        fd = open(file_name, 'rb')
        message.ParseFromString(fd.read())
    except Exception, e:
        print str(e)
        quit()

######################################################################

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print '''
            Usage:
                converter2.py <excel> <protocol> [store path, default .]
            Example:
                converter2.py attr.xlsx attr.proto ../output
        '''
        quit()

    excel = sys.argv[1]
    proto = sys.argv[2]
    output = len(sys.argv) > 3 and sys.argv[3] or '.'
    convert(excel, proto, output)

