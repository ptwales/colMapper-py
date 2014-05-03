import xlmp
import xlrd
import string
import unittest

DEFAULT_LIMIT = 26
DEFAULT_RANGE = list(range(0, DEFAULT_LIMIT))
XL_EXTS = ['.xls', '.xlsx'] #, '.ods'] ods not supported...
TEST_NAME = 'test'
TEST_BOOKS = [TEST_NAME + exts for exts in XL_EXTS]
TEST_MATRIX = [[unicode(string.uppercase[c] + str(r + 1), 'utf8') 
                for c in DEFAULT_RANGE] for r in DEFAULT_RANGE]
TRANSPOSED = [[unicode(string.uppercase[c] + str(r + 1), 'utf8')
               for r in DEFAULT_RANGE] for c in DEFAULT_RANGE]
SUB_MATRIX = [row[3:15] for row in TEST_MATRIX[5:20]]


def get_test_sheet(book_path):
    return xlrd.open_workbook(book_path).sheet_by_index(0)


def loop_range(func, *args):
    for initial in range(0, DEFAULT_LIMIT - 1):
        for final in range(initial + 1, DEFAULT_LIMIT):
            func(initial, final, *args)


class TestIxl(unittest.TestCase):

    def setUp(self):
        self.xl_interface = xlmp._ExcelInterface(True)

    def test_transpose(self):
        self.assertEqual(TEST_MATRIX,
            self.xl_interface.__ExcelInterface__transpose(TEST_MATRIX)
            )
        self.xl_interface.by_row = False
        self.assertEqual(TRANSPOSED,
            self.xl_interface.__ExcelInterface__transpose(TEST_MATRIX)
            )

    def test_read_sheet(self):

        def scan_ranges(i, f, s):
            self.xlrw.by_row = True
            self.assertEqual(TEST_MATRIX[i:f],
                self.xl_interface.read_sheet(s, r0=i, rf=f)
                )
            self.assertEqual([m[i:f] for m in TEST_MATRIX],
                self.xlrw.read_sheet(s, c0=i, cf=f)
                )
            self.xlrw.by_row = False
            self.assertEqual(TRANSPOSED[i:f],
                self.xl_interface.read_sheet(s, c0=i, cf=f)
                )
            self.assertEqual([m[i:f] for m in TRANSPOSED],
                self.xl_interface.read_sheet(s, r0=i, rf=f)
                )

        sheets = [get_test_sheet(book) for book in TEST_BOOKS]
        for a_sheet in sheets:
            for initial in DEFAULT_RANGE:
                for final in range(initial + 1, DEFAULT_LIMIT):
                    scan_ranges(initial, final, a_sheet)

    def test_read_book(self):
        for book in TEST_BOOKS:
            self.assertEqual(TEST_MATRIX,
                self.xl_interface.read_book(book, 0)
                )
            self.assertEqual(TEST_MATRIX,
                self.xl_interface.read_book(book, TEST_NAME)
                )
            # I have already tested argument lists
            self.assertEqual(SUB_MATRIX,
                self.xl_interface.read_book(book, 0, r0=5, rf=20, c0=3, cf=15)
                )

    def test_guess_read(self):
        for book in TEST_BOOKS:
            self.assertEqual(TEST_MATRIX,
                self.xl_interface.guess_read(
                    sheet=xlrd.open_workbook(book).sheet_by_index(0))
                )
            self.assertEqual(TEST_MATRIX,
                self.xl_interface.guess_read(book_path=book)
                )

    #def test_write_sheet(self):
    #def test_write_book(self):
    #def test_guess_write(self):
    
    
if __name__ == '__main__':
    unittest.main()
