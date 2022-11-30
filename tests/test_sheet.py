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
        ['Qué?', 'Qué?', 'Qué?', 'Quién?', 'Para quién?', 'Para quién?', 'Dónde?', 'Cuándo?', 'Cuándo?'],
        ['Registro', 'Sector/Cluster', 'Subsector', 'Organización', 'Hombres', 'Mujeres', 'País', 'Departamento/Provincia/Estado', 'Departamento/Provincia/Estado'],
        ['Departamento/Provincia/Estado', '#sector+es', '#subsector+es', '#org+es', '#targeted+f', '#targeted+m', '#country', '#adm1', '#date+reported'],
        ['001', 'WASH', 'Higiene', 'ACNUR', '100', '100', 'Panamá', 'Los Santos', '1 March 2015'],
        ['002', 'Salud', 'Vacunación', 'OMS', 'Colombia', 'Cauca'],
        ['003', 'Educación', 'Formación de enseñadores', 'UNICEF', '250', '300', 'Colombia', 'Chocó'],
        ['004', 'WASH', 'Urbano', 'OMS', '80', '95', 'Venezuela', 'Amazonas'],
    ]
    
    def setUp(self):
        self.workbook = xlsxr.Workbook(filename=resolve_path("simple.xlsx"))
        self.sheet = self.workbook.get_sheet(1)

    def test_name(self):
        self.assertEqual('input-valid', self.sheet.name)

    def test_sheetId(self):
        self.assertEqual('1', self.sheet.sheetId)

    def test_state(self):
        self.assertEqual('visible', self.sheet.state)

    def test_rel_id(self):
        self.assertEqual('rId2', self.sheet.rel_id)

    def test_target(self):
        self.assertEqual('worksheets/sheet1.xml', self.sheet.target)

    def test_rows(self):
        content = [row for row in self.sheet.rows()]
        self.assertEqual(self.EXPECTED_CONTENT, content)


        
