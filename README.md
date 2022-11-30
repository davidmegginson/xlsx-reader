xlsx-reader
===========

Python3 library optimised for reading very large Excel XLSX files, including those with hundreds of columns as well as rows. For now, everything is a string.

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


# License

This is free and unencumbered software released into the public domain. See UNLICENSE.md for details.