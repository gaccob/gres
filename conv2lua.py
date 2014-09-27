import xdrlib ,sys
import xlrd
import getopt

##################################################################

def convert_sheet(sheet, output_path):
    nrows = sheet.nrows
    ncols = sheet.ncols

    keys = {}
    for i in range(0, ncols):
        keys[i] = sheet.cell(0, i).value

    output = []
    output.append("Config = {")
    for i in range(1, nrows):
        output.append("    {")
        for j in keys:
            val = sheet.cell(i, j).value
            if val == '' or val == u'':
                continue
            line = "        %s = %s" % (keys[j], val)
            line = ((j == ncols - 1) and line or ("%s, " % line)).encode("utf-8")
            output.append(line)
        output.append((i == nrows - 1) and "    }" or "    },")
    output.append("}")

    name = '%s/%s.lua' % (output_path, sheet.name)
    try:
        fd = open(name, 'w')
    except Exception, e:
        print str(e)
        return

    output = [line + '\n' for line in output]
    fd.writelines(output)
    fd.close()
    print 'convert %s success' % name

##################################################################

def convert(xls, output_path):
    try:
        book = xlrd.open_workbook(xls)
    except Exception, e:
        print str(e)
        return

    # get sheet one by one
    for sheet in book.sheets():
        convert_sheet(sheet, output_path)

##################################################################

def usage():
    print '''
    conv2lua.py usage:
        -f, --excel_file    <excel file name>               [required]
        -O, --output_path   <convert result's path>         [optional]
    '''
    quit()

##################################################################

if __name__=="__main__":

    sopt = 'f:O:'
    lopt = ['excel_file=', 'output_path=']
    try:
        opts, args = getopt.getopt(sys.argv[1:], sopt, lopt)
    except Exception, e:
        print str(e) 
        usage()

    excel = ''
    output_path = '.'
    for opt, val in opts:
        if opt in ('-f', '--excel_file'):
            excel = val
        elif opt in ('-O', '--output_path'):
            output_path = val

    if excel == '':
        usage()

    convert(excel, output_path)

