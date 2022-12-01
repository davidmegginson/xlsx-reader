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

    EXPECTED_CONTENT = [
        ['Qué?', None, None, 'Quién?', 'Para quién?', None, 'Dónde?', None, 'Cuándo?'],
        ['Registro', 'Sector/Cluster', 'Subsector', 'Organización', 'Hombres', 'Mujeres', 'País', 'Departamento/Provincia/Estado', None],
        [None, '#sector+es', '#subsector+es', '#org+es', '#targeted+f', '#targeted+m', '#country', '#adm1', '#date+reported'],
        ['001', 'WASH', 'Higiene', 'ACNUR', '100', '100', 'Panamá', 'Los Santos', '42064'], # FIXME - is a date
        ['002', 'Salud', 'Vacunación', 'OMS', 'Colombia', 'Cauca'],
        ['003', 'Educación', 'Formación de enseñadores', 'UNICEF', '250', '300', 'Colombia', 'Chocó'],
        ['004', 'WASH', 'Urbano', 'OMS', '80', '95', 'Venezuela', 'Amazonas'],
    ]
    
    EXPECTED_CONTENT_CONVERTED = [
        ['Qué?', None, None, 'Quién?', 'Para quién?', None, 'Dónde?', None, 'Cuándo?'],
        ['Registro', 'Sector/Cluster', 'Subsector', 'Organización', 'Hombres', 'Mujeres', 'País', 'Departamento/Provincia/Estado', None],
        [None, '#sector+es', '#subsector+es', '#org+es', '#targeted+f', '#targeted+m', '#country', '#adm1', '#date+reported'],
        ['001', 'WASH', 'Higiene', 'ACNUR', 100, 100, 'Panamá', 'Los Santos', 42064], # FIXME - is a date
        ['002', 'Salud', 'Vacunación', 'OMS', 'Colombia', 'Cauca'],
        ['003', 'Educación', 'Formación de enseñadores', 'UNICEF', 250, 300, 'Colombia', 'Chocó'],
        ['004', 'WASH', 'Urbano', 'OMS', 80, 95, 'Venezuela', 'Amazonas'],
    ]
    
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

    def test_rows(self):
        content = [row for row in self.sheet.rows]
        print(content)
        self.assertEqual(self.EXPECTED_CONTENT, content)

    def test_rows_converted(self):
        self.workbook.convert_values = True
        content = [row for row in self.sheet.rows]
        self.assertEqual(self.EXPECTED_CONTENT_CONVERTED, content)


        
