""" Unit tests for the xlrx module
Started by David Megginson, 2020-03-20

@author: David Megginson
@organization: UN Centre for Humanitarian Data
@license: Public Domain
@date: Started 2020-03-20
"""

import unittest
import xlsxr

from . import resolve_path

class TestWorkbook(unittest.TestCase):

    def test_open_workbook_from_filename(self):
        xlsxr.Workbook(filename=resolve_path("simple.xlsx"))

    def test_open_workbook_from_stream(self):
        with open(resolve_path("simple.xlsx"), "rb") as input:
            xlsxr.Workbook(stream=input)

    def test_open_workbook_from_url(self):
        xlsxr.Workbook(url="https://github.com/davidmegginson/xlsx-reader/blob/master/tests/files/simple.xlsx?raw=true")

    def test_open_non_excel_archive(self):
        with self.assertRaises(TypeError):
            xlsxr.Workbook(filename=resolve_path("not-excel.zip"))

    def test_sheet_count(self):
        workbook = xlsxr.Workbook(filename=resolve_path("simple.xlsx"))
        self.assertEqual(1, workbook.sheet_count)
