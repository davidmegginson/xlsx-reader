""" Unit tests for the xlrx module

@author: David Megginson
@organization: UN Centre for Humanitarian Data
@license: Public Domain
@date: Started 2020-03-20
"""

import unittest
import xlsxr

from . import resolve_path

class TestWorkbook(unittest.TestCase):

    def setUp(self):
        self.workbook = xlsxr.Workbook(filename=resolve_path("simple.xlsx"))


    def test_open_workbook_from_stream(self):
        with open(resolve_path("simple.xlsx"), "rb") as input:
            xlsxr.Workbook(stream=input)

    def test_open_workbook_from_url(self):
        xlsxr.Workbook(url="https://github.com/davidmegginson/xlsx-reader/blob/main/tests/files/simple.xlsx?raw=true")

    def test_open_non_excel_archive(self):
        with self.assertRaises(TypeError):
            xlsxr.Workbook(filename=resolve_path("not-excel.zip"))

    def test_sheet_count(self):
        self.assertEqual(1, self.workbook.sheet_count)

    def test_shared_strings(self):
        self.assertTrue('UNICEF' in self.workbook.shared_strings)
        self.assertTrue('Sector/Cluster' in self.workbook.shared_strings)

    def test_rels(self):
        self.assertTrue('rId2' in self.workbook.rels)

    def test_get_sheet(self):
        self.assertIsNotNone(self.workbook.get_sheet(1))

    def test_get_bad_sheet(self):
        with self.assertRaises(IndexError):
            self.workbook.get_sheet(-1)
        with self.assertRaises(IndexError):
            self.workbook.get_sheet(7)
        
