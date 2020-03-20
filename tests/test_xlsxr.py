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

    def test_open_workbook_filename(self):
        filename = resolve_path("simple.xlsx")
        xlsxr.Workbook(filename=filename)
        self.assertTrue(True)
