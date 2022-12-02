xlsx-reader
===========

Python3 library optimised for reading very large Excel XLSX files, including those with hundreds of columns as well as rows. 

# Simple example

```
from xlsxr import Workbook

workbook = Workbook(filename="myworkbook.xlsx", convert_values=True)

for sheet in workbook.sheets:
    print("Sheet ", sheet.name)
    for row in sheet.rows:
        print(row)
```

# Conversions

By default, everything is a string, and all dates and datetimes will appear in ISO 8601 format (YYYY-mm-dd or YYYY-mm-ddTHH:MM:SS). If you supply the option convert_values to the Worksheet constructor, the library will convert numbers to ints or floats, and dates to datetime.datetime or datetime.date objects. There is no attempt to handle standalone times.


# License

This is free and unencumbered software released into the public domain. See UNLICENSE.md for details.