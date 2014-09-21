import xdrlib ,sys
import xlrd

def convert(xls, sheet, dest):
    try:
        book = xlrd.open_workbook(xls)
    except Exception, e:
        print str(e)
        return

    table = book.sheet_by_name(sheet)
    nrows = table.nrows
    ncols = table.ncols

    keys = {}
    for i in range(0, ncols):
        keys[i] = table.cell(0, i).value

    output = []
    output.append("data = {")
    for i in range(1, nrows):
        output.append("    {")
        for j in keys:
            line = "        %s = %s" % (keys[j], table.cell(i, j).value)
            line = ((j == ncols - 1) and line or ("%s, " % line)).encode("utf-8")
            output.append(line)
        output.append((i == nrows - 1) and "    }" or "    },")
    output.append("}")

    try:
        fd = open(dest, "w")
    except Exception, e:
        print str(e)
        return
    output = [line + '\n' for line in output]
    fd.writelines(output)
    fd.close()

if __name__=="__main__":
    convert("./sample/prop.xls", "prop", "./sample/prop.lua")

