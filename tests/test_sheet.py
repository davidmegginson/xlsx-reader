""" Unit tests for the xlrx.sheet module

@author: David Megginson
@organization: UN Centre for Humanitarian Data
@license: Public Domain
@date: Started 2020-03-20
"""

import unittest
import datetime, xlsxr

from . import resolve_path

class TestSheet(unittest.TestCase):

    EXPECTED_COLS = [
        {'collapsed': False, 'hidden': False, 'min': 1, 'max': 1, 'style': '0'},
        {'collapsed': False, 'hidden': False, 'min': 2, 'max': 2, 'style': '0'},
        {'collapsed': False, 'hidden': False, 'min': 3, 'max': 3, 'style': '0'},
        {'collapsed': False, 'hidden': False, 'min': 4, 'max': 4, 'style': '0'},
        {'collapsed': False, 'hidden': False, 'min': 5, 'max': 5, 'style': '0'},
        {'collapsed': False, 'hidden': False, 'min': 6, 'max': 6, 'style': '0'},
        {'collapsed': False, 'hidden': False, 'min': 7, 'max': 7, 'style': '0'},
        {'collapsed': False, 'hidden': False, 'min': 8, 'max': 8, 'style': '0'},
        {'collapsed': False, 'hidden': False, 'min': 9, 'max': 9, 'style': '0'},
    ]

    EXPECTED_ROWS = [
        ['Qué?', '', '', 'Quién?', 'Para quién?', '', 'Dónde?', '', 'Cuándo?'],
        ['Registro', 'Sector/Cluster', 'Subsector', 'Organización', 'Hombres', 'Mujeres', 'País', 'Departamento/Provincia/Estado', ''],
        ['', '#sector+es', '#subsector+es', '#org+es', '#targeted+f', '#targeted+m', '#country', '#adm1', '#date+reported'],
        ['001', 'WASH', 'Higiene', 'ACNUR', '100', '100', 'Panamá', 'Los Santos', '2015-03-01'], # FIXME - is a date
        ['002', 'Salud', 'Vacunación', 'OMS', '', '', 'Colombia', 'Cauca'],
        ['003', 'Educación', 'Formación de enseñadores', 'UNICEF', '250', '300', 'Colombia', 'Chocó'],
        [],
        ['004', 'WASH', 'Urbano', 'OMS', '80', '95', 'Venezuela', 'Amazonas'],
    ]
    
    EXPECTED_ROWS_CONVERTED = [
        ['Qué?', '', '', 'Quién?', 'Para quién?', '', 'Dónde?', '', 'Cuándo?'],
        ['Registro', 'Sector/Cluster', 'Subsector', 'Organización', 'Hombres', 'Mujeres', 'País', 'Departamento/Provincia/Estado', ''],
        ['', '#sector+es', '#subsector+es', '#org+es', '#targeted+f', '#targeted+m', '#country', '#adm1', '#date+reported'],
        ['001', 'WASH', 'Higiene', 'ACNUR', 100, 100, 'Panamá', 'Los Santos', datetime.date(2015,3,1)], # FIXME - is a date
        ['002', 'Salud', 'Vacunación', 'OMS', '', '', 'Colombia', 'Cauca'],
        ['003', 'Educación', 'Formación de enseñadores', 'UNICEF', 250, 300, 'Colombia', 'Chocó'],
        [],
        ['004', 'WASH', 'Urbano', 'OMS', 80, 95, 'Venezuela', 'Amazonas'],
    ]

    EXPECTED_MERGES = ['A1:C1', 'E1:F1', 'G1:H1',]
    
    def setUp(self):
        self.workbook = xlsxr.Workbook(filename=resolve_path("simple.xlsx"))
        self.sheet = self.workbook.sheets[0]

    def test_name(self):
        self.assertEqual('input-valid', self.sheet.name)

    def test_sheet_id(self):
        self.assertEqual('1', self.sheet.sheet_id)

    def test_state(self):
        self.assertEqual('visible', self.sheet.state)

    def test_relation_id(self):
        self.assertEqual('rId2', self.sheet.relation_id)

    def test_filename(self):
        self.assertEqual('xl/worksheets/sheet1.xml', self.sheet.filename)

    def test_cols(self):
        self.assertEqual(self.EXPECTED_COLS, self.sheet.cols)

    def test_rows(self):
        self.assertEqual(self.EXPECTED_ROWS, self.sheet.rows)

    def test_rows_converted(self):
        self.workbook.convert_values = True
        self.assertEqual(self.EXPECTED_ROWS_CONVERTED, self.sheet.rows)

    def test_merges(self):
        self.assertEqual(self.EXPECTED_MERGES, self.sheet.merges)

    def test_get_col(self):
        self.assertEqual(self.EXPECTED_COLS[2], self.sheet.get_col(2))

        
