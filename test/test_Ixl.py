from xlmp import Ixl
import xlrd
import string
import unittest

DEFAULT_RANGE = list(range(0, 26))
XL_EXTS = ['.xls', '.xlsx', '.ods']
TEST_NAME = 'test'
M = [[string.uppercase[c] + str(r + 1) for r in DEFAULT_RANGE] for c in DEFAULT_RANGE]
MT = [[string.uppercase[c] + str(r + 1) for c in DEFAULT_RANGE] for r in DEFAULT_RANGE]


def get_test_sheet(self, book_path):
        return xlrd.open_workbook(TEST_NAME + e).sheet_by_index(0)


class TestIxl(unittest.TestCase):
    
    def setUp(self):
        self.xlwt = Ixl()
    
    def test_transpose(self):
        self.assertEqual(M, self.xlwt.__transpose(M))
        self.xlwt.by_row = False
        self.assertEqual(MT, self.xlwt.__transpose(M))
        
    def test_read_sheet(self):
        sheets = [get_test_sheet(TEST_NAME + e) for e in XL_EXTS]
        # TODO: wrap this to test ranges
        for s in sheets:
            self.xlwt.by_row = True
            self.assertEqual(M, self.xlwt.read_sheet(s))
            self.xlwt.by_row = False
            self.assertEqual(MT, self.xlwt.read_sheet(s))
    
    
if __name__ == '__main__':
    unittest.main()
