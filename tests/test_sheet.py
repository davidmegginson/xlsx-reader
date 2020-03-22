""" Unit tests for the xlrx.sheet module

@author: David Megginson
@organization: UN Centre for Humanitarian Data
@license: Public Domain
@date: Started 2020-03-20
"""

import unittest
import xlsxr

from . import resolve_path

class TestSheet(unittest.TestCase):

    def setUp(self):
        self.workbook = xlsxr.Workbook(filename=resolve_path("simple.xlsx"))
        self.sheet = self.workbook.get_sheet(1)

    def test_placeholder(self):
        pass
