xlsx-reader
===========

Python3 library optimised for reading very large Excel XLSX files, including those with hundreds of columns as well as rows. 

# Simple example

```
from xlsxr import Workbook

with open('myworkbook.xlsx', 'rb') as input:
    workbook = Workbook(stream=input)

    for n in range(1, workbook.sheet_count + 1):
        sheet = workbook.get_sheet(n)
        print("Sheet ", sheet.name)
        for row in sheet.rows:
            print("    ", row)
```

# Conversions

By default, everything is a string. If you supply the optional convert_values to the Worksheet constructor, the library will convert numbers. For now, dates are just weird numbers (Excel doesn't flag dates as dates per se; you have to figure it it from the style template).


# License

This is free and unencumbered software released into the public domain. See UNLICENSE.md for details.