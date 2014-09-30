import xlmp
import xlrd
import xlwt
import csv
import string
import unittest

XL_EXTS = ['.xls', '.xlsx']  # , '.ods'] ods not supported...
Test_Matrix = [[str(string.ascii_uppercase[c] + str(r + 1))
    for c in range(0,26)] for r in range(0,26)]
C0, CF, R0, RF = 3, 15, 5, 20
Sub_Matrix = [row[C0:CF] for row in TestMatrix[R0:RF]]



class TestSheetIO(unittest.TestCase):

    def setUp(self):
        pass

    def test_factory(self):
        pass

    def test_transpose(self):
        pass

class Testxlrd(unittest.TestCase):

    def setUp(self):
        pass

    def test_read_sheet(self):
        pass

    def test_read_subsection(self):
        pass

class Testxlwt(unittest.TestCase):
    
    def setUp(self):
        pass

    def test_write_sheet(self):
        pass

    def test_write_subsection(self):
        pass

class Testcsv(unittest.TestCase):

    def setUp(self):
        pass

    def test_read_sheet(self):
        pass

    def test_read_subsection(self):
        pass

    def test_write_sheet(self):
        pass

    
if __name__ == '__main__':
    unittest.main()
