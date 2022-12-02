""" Unit tests for the xlrx.style module

@author: David Megginson
@organization: UN Centre for Humanitarian Data
@license: Public Domain
@date: Started 2020-03-20
"""

import unittest
import xlsxr

from . import resolve_path

class TestStyles(unittest.TestCase):

    def setUp(self):
        self.workbook = xlsxr.Workbook(filename=resolve_path("simple.xlsx"))
        self.styles = self.workbook.styles

    def test_number_formats(self):
        self.assertEqual(EXPECTED_NUMBER_FORMATS, self.styles.number_formats)

    def test_cell_style_formats(self):
        self.assertEqual(EXPECTED_CELL_STYLE_FORMATS, self.styles.cell_style_formats)

    def test_cell_formats(self):
        print(self.styles.cell_formats)
        self.assertEqual(EXPECTED_CELL_FORMATS, self.styles.cell_formats)

    def test_cell_styles(self):
        self.assertEqual(EXPECTED_CELL_STYLES, self.styles.cell_styles)


#
# Test data
#

EXPECTED_NUMBER_FORMATS = {
    '164': 'General',
    '165': 'mmm\\ d", "yyyy',
}

EXPECTED_CELL_STYLE_FORMATS = [
    {'numFmtId': '164', 'applyProtection': True, 'protection': {'locked': True, 'hidden': False}},
    {'numFmtId': '0', 'applyProtection': False},
    {'numFmtId': '0', 'applyProtection': False},
    {'numFmtId': '0', 'applyProtection': False},
    {'numFmtId': '0', 'applyProtection': False},
    {'numFmtId': '0', 'applyProtection': False},
    {'numFmtId': '0', 'applyProtection': False},
    {'numFmtId': '0', 'applyProtection': False},
    {'numFmtId': '0', 'applyProtection': False},
    {'numFmtId': '0', 'applyProtection': False},
    {'numFmtId': '0', 'applyProtection': False},
    {'numFmtId': '0', 'applyProtection': False},
    {'numFmtId': '0', 'applyProtection': False},
    {'numFmtId': '0', 'applyProtection': False},
    {'numFmtId': '0', 'applyProtection': False},
    {'numFmtId': '43', 'applyProtection': False},
    {'numFmtId': '41', 'applyProtection': False},
    {'numFmtId': '44', 'applyProtection': False},
    {'numFmtId': '42', 'applyProtection': False},
    {'numFmtId': '9', 'applyProtection': False},
]

EXPECTED_CELL_FORMATS = [
    {'numFmtId': '164', 'applyProtection': False, 'protection': {'locked': True, 'hidden': False}, 'has_date': False, 'has_time': False,},
    {'numFmtId': '164', 'applyProtection': False, 'protection': {'locked': True, 'hidden': False}, 'has_date': False, 'has_time': False,},
    {'numFmtId': '164', 'applyProtection': False, 'protection': {'locked': True, 'hidden': False}, 'has_date': False, 'has_time': False,},
    {'numFmtId': '164', 'applyProtection': False, 'protection': {'locked': True, 'hidden': False}, 'has_date': False, 'has_time': False,},
    {'numFmtId': '165', 'applyProtection': False, 'protection': {'locked': True, 'hidden': False}, 'has_date': True, 'has_time': False,},
]

EXPECTED_CELL_STYLES = [
    {'name': 'Normal', 'xfId': '0', 'builtinId': '0'},
    {'name': 'Comma', 'xfId': '15', 'builtinId': '3'},
    {'name': 'Comma [0]', 'xfId': '16', 'builtinId': '6'},
    {'name': 'Currency', 'xfId': '17', 'builtinId': '4'},
    {'name': 'Currency [0]', 'xfId': '18', 'builtinId': '7'},
    {'name': 'Percent', 'xfId': '19', 'builtinId': '5'},
]
