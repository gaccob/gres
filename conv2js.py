import sys,getopt
import xlrd

##################################################################

def output_cell(output, shift, dict, sheet, row):
    for (k,v) in dict.items():
        if type(v) == type({}):
            output.append("\t" * shift + "\"" + k + "\": {")
            output_cell(output, shift + 1, v, sheet, row)
            output.append("\t" * shift + "},");
        else:
            val = sheet.cell(row, v).value
            if type(val) == type(1) or type(val) == type(1.1):
                output.append("\t" * shift + "\"" + k + "\": " + str(val) + ",")
            else:
                val = unicode(val, 'utf8')
                output.append("\t" * shift + "\"" + k + "\": \"" + val + "\",")

def convert_sheet(sheet, output_path):
    nrows = sheet.nrows
    ncols = sheet.ncols

    keys = {}
    for i in range(0, ncols):
        parent = keys
        fields = (sheet.cell(0, i).value).split('.')
        for j in range(0, len(fields)):
            field = fields[j]
            if j == len(fields) - 1:
                parent[field] = i
            else:
                if parent.get(field) == None:
                    parent[field] = {}
                parent = parent[field]

    output = []
    output.append("var " + sheet.name + " = {")
    for row in range(1, nrows):
        output.append("\t\"" + str(row - 1) + "\": {");
        output_cell(output, 2, keys, sheet, row);
        output.append("\t},");
    output.append("};")
    output.append("module.exports = " + sheet.name + ";");

    name = '%s/%s.js' % (output_path, sheet.name)
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

