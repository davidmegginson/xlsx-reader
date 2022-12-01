import csv, io, re, sys, xlsxr

def convert_xlsx(source, index=0):
    """ Convert a sheet of an Excel workbook to CSV """

    # Figure out if it's a pipe, a filename, or a URL
    if hasattr(source, 'seek'):
        workbook = xlsxr.Workbook(stream=source)
    elif re.match('^https?:', source):
        workbook = xlsxr.Workbook(url=source)
    else:
        workbook = xlsxr.Workbook(filename=source)

    output = csv.writer(sys.stdout)

    for row in workbook.sheets[index].rows:
        output.writerow(row)

if len(sys.argv) == 1:
    convert_xlsx(io.BytesIO(sys.stdin.buffer.read()))
elif len(sys.argv) == 2:
    convert_xlsx(sys.argv[1])
else:
    print("Usage: python3 {} [file-or-url?]".format(sys.argv[0]), file=sys.stderr)
