xlsx-reader
===========

Python3 library optimised for reading very large Excel XLSX files, including those with hundreds of columns as well as rows. 

## Simple example

```
from xlsxr import Workbook

workbook = Workbook(filename="myworkbook.xlsx", convert_values=True)

for sheet in workbook.sheets:
    print("Sheet ", sheet.name)
    for row in sheet.rows:
        print(row)
```

## Conversions

By default, everything is a string, and all dates and datetimes will appear in ISO 8601 format (YYYY-mm-dd or YYYY-mm-ddTHH:MM:SS). If you supply the option _convert\_values_ to the Worksheet constructor, the library will convert numbers to ints or floats, and dates to datetime.datetime or datetime.date objects. There is no attempt to handle standalone times.

Empty cells appear as the empty string ''.

## xlsxr.workbook.Workbook class

### Constructor

The xlsxr.Workbook class constructor takes the following keyword arguments:

Argument | Description
-- | --
filename | Path to an Excel file on the local filesystem.
stream | A file-like object (byte stream)
url | The URL of a remote Excel file
convert\_values | If True, convert numbers and dates from strings to Python values (default is False)

You may specify only one of _filename,_ _stream,_ or _url._

### Properties

Property | Description
-- | --
sheets | A list of xlsxr.sheet.Sheet objects
styles | A list of xlsxr.style.Style objects

## xlsxr.sheet.Sheet class

### Properties

Property | Description
workbook | The parent Workbook objet
name | The name of the sheet
sheet\_id | The internal identifier of the sheet
state | The state of the sheet (normally 'visible' or 'hidden')
relation\_id | ??
cols | A list of metadata for each column.
rows | A list of the data rows in the sheet (parsed on demand).
merges | A list of merges in the sheet (parsed on demand).

Each row is a list of scalar values. The will all be strings or None unless you specified the _convert\_values_ option for the Workbook.

Merges appear as strings defining ranges, e.g. "A1:C3".

### Columns

Columns are represented as dict objects with the following properties:

Key | Description
-- | --
collapsed | True if the column is collapsed.
hidden | True if the column is hidden
min | ??
max | ??
style | A key into the _styles_ property of the workbook.

## xlsxr.style.Style class

### Properties

Property | Description
-- | --
number\_formats | ??
cell\_style\_formats | ??
cell\_formats | ??
cell\_styles | ??

# License

This is free and unencumbered software released into the public domain. See UNLICENSE.md for details.