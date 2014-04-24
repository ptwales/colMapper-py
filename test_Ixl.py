import xlmp
import xlrd
import string
import unittest

DEFAULT_LIMIT = 26
DEFAULT_RANGE = list(range(0, DEFAULT_LIMIT))
XL_EXTS = ['.xls', '.xlsx'] #, '.ods'] ods not supported...
TEST_NAME = 'test'
M = [[unicode(string.uppercase[c] + str(r + 1), 'utf8') for c in DEFAULT_RANGE] for r in DEFAULT_RANGE]
MT = [[unicode(string.uppercase[c] + str(r + 1), 'utf8') for r in DEFAULT_RANGE] for c in DEFAULT_RANGE]


def get_test_sheet(book_path):
    return xlrd.open_workbook(book_path).sheet_by_index(0)

def loop_range(func, *args):
    for i in range(0, DEFAULT_LIMIT - 1):
        for f in range(i + 1, DEFAULT_LIMIT):
            func(i, f, *args)

class TestIxl(unittest.TestCase):
    
    def setUp(self):
        self.xlrw = xlmp.Ixl(True)
    
    def test_transpose(self):
        self.assertEqual(M, self.xlrw.__transpose(M))
        self.xlrw.by_row = False
        self.assertEqual(MT, self.xlrw.__transpose(M))

    # TODO: test Slicing  
    def test_read_sheet(self):
        
        def test_case(i, f, s):
            self.xlrw.by_row = True
            self.assertEqual(M[i:f + 1], self.xlrw.read_sheet(s, r0=i, rf=f))
            self.assertEqual([m[i:f + 1] for m in M], self.xlrw.read_sheet(s, c0=i, cf=f))
            self.xlrw.by_row = False
            self.assertEqual(MT[i:f + 1], self.xlrw.read_sheet(s, c0=i, cf=f))
            self.assertEqual([m[i:f + 1] for m in MT], self.xlrw.read_sheet(s, r0=i, rf=f))
            
        sheets = [get_test_sheet(TEST_NAME + e) for e in XL_EXTS]  
        for s in sheets:
            self.xlrw.by_row = True
            self.assertEqual(M, self.xlrw.read_sheet(s))
            self.xlrw.by_row = False
            self.assertEqual(MT, self.xlrw.read_sheet(s))
            #loop_range(test_case, s)
            
                    
    #def test_read_book(self):
    #def test_guess_read(self):
    #def test_write_sheet(self):
    #def test_write_book(self):
    #def test_guess_write(self):
    
    
if __name__ == '__main__':
    unittest.main()
