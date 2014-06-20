import xlmp
import xlrd
import xlwt
import string
import unittest

DEFAULT_LIMIT = 26
DEFAULT_RANGE = list(range(0, DEFAULT_LIMIT))
XL_EXTS = ['.xls', '.xlsx']  # , '.ods'] ods not supported...
TEST_NAME = 'test'
TEST_BOOKS = [TEST_NAME + exts for exts in XL_EXTS]
TEST_MATRIX = [[unicode(string.ascii_uppercase[c] + str(r + 1), 'utf8')
                for c in DEFAULT_RANGE] for r in DEFAULT_RANGE]
TRANSPOSED = [[unicode(string.ascii_uppercase[c] + str(r + 1), 'utf8')
               for r in DEFAULT_RANGE] for c in DEFAULT_RANGE]
C0, CF, R0, RF = 3, 15, 5, 20
SUB_MATRIX = [row[3:15] for row in TEST_MATRIX[5:20]]


def get_test_sheet(book_path):
    return xlrd.open_workbook(book_path).sheet_by_index(0)


def loop_range(func, *args):
    for initial in range(0, DEFAULT_LIMIT - 1):
        for final in range(initial + 1, DEFAULT_LIMIT):
            func(initial, final, *args)


class TestExcelInterface(unittest.TestCase):

    def setUp(self):
        self.xl_interface = xlmp._ExcelInterface(True)

    def test_transpose(self):
        self.assertEqual(TEST_MATRIX,
            self.xl_interface._ExcelInterface__transpose(TEST_MATRIX)
            )
        self.xl_interface.by_row = False
        self.assertEqual(TRANSPOSED,
            self.xl_interface._ExcelInterface__transpose(TEST_MATRIX)
            )

    def test_read_sheet(self):

        def scan_ranges(i, f, s):
            self.xl_interface.by_row = True
            self.assertEqual(TEST_MATRIX[i:f],
                self.xl_interface.read_sheet(s, r0=i, rf=f)
                )
            self.assertEqual([m[i:f] for m in TEST_MATRIX],
                self.xl_interface.read_sheet(s, c0=i, cf=f)
                )
            self.xl_interface.by_row = False
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
                self.xl_interface.read_book(book, 0, r0=R0, rf=RF, c0=C0, cf=CF)
                )

    def test_guess_read(self):
        for book in TEST_BOOKS:
            self.assertEqual(TEST_MATRIX,
                self.xl_interface.guess_read(
                    sheet=xlrd.open_workbook(book).sheet_by_index(0)
                    )
                )
            self.assertEqual(TEST_MATRIX,
                self.xl_interface.guess_read(book_path=book)
                )

    def test_write_sheet(self):
        book_path = TEST_BOOKS[0]
        book = xlwt.Workbook()
        sheet = book.add_sheet('xlmp')
        self.xl_interface.write_sheet(TEST_MATRIX, sheet)
        book.save("write_" + book_path)

    def test_write_sheet_sub(self):
        book_path = TEST_BOOKS[0]
        book = xlwt.Workbook()
        sheet = book.add_sheet('xlmp')
        self.xl_interface.write_sheet(SUB_MATRIX, sheet, r0=R0, c0=C0)
        book.save("write_sub_" + book_path)

    def test_write_sheet_trans(self):
        book_path = TEST_BOOKS[0]
        self.xl_interface.by_row = False
        book = xlwt.Workbook()
        sheet = book.add_sheet('xlmp')
        self.xl_interface.write_sheet(TEST_MATRIX, sheet)
        book.save("write_trans_" + book_path)

    #def test_write_book(self):
    #def test_guess_write(self):

i_INDEX_COMMAND = {i: i for i in DEFAULT_RANGE}
i_NAME_KEY_COMMAND = {string.ascii_uppercase[i]: i for i in DEFAULT_RANGE}
i_NAME_VAL_COMMAND = {i: string.ascii_uppercase[i] for i in DEFAULT_RANGE}
i_OFFSET_COMMAND = {i + 1: i + 1 for i in DEFAULT_RANGE}

o_INDEX_COMMAND = {i: (lambda row, i=i: row[i]) for i in DEFAULT_RANGE}

TEST_ROW = TEST_MATRIX[15]
TEST_STRING = "THISISANEXPARROT"
TEST_INT_ROW = [1, -1600, 4999, -5000, 4000]
PARROT_INDEXES = [19, 7, 8, 18,
                  8, 18, 0, 13,
                  4, 23,
                  15, 0, 17, 17, 14, 19]


class TestmpCmd(unittest.TestCase):

    #
    # KEY TESTING
    #
    # stupid test case; nothing should change
    # index -> index
    def test_index_key(self):
        cmd = xlmp.mpCmd(i_INDEX_COMMAND)
        self.assertEqual(cmd.keys(), o_INDEX_COMMAND.keys())

    # index -> index - offset
    def test_index_key_offset(self):
        cmd = xlmp.mpCmd(i_OFFSET_COMMAND, offset=1)
        self.assertEqual(cmd.keys(), o_INDEX_COMMAND.keys())

    # name -> index
    def test_name_key_replace(self):
        cmd = xlmp.mpCmd(i_NAME_KEY_COMMAND)
        self.assertEqual(cmd.keys(), o_INDEX_COMMAND.keys())

    # there is no name offset
    # name -> index
    def test_name_key_offset(self):
        cmd = xlmp.mpCmd(i_NAME_KEY_COMMAND, offset=1)
        self.assertEqual(cmd.keys(), o_INDEX_COMMAND.keys())

    # keys can only be int or str
    def test_invalid_key(self):
        with self.assertRaises(TypeError):
            cmd = xlmp.mpCmd({[1, 2, 3]: 4})
            cmd = xlmp.mpCmd({2.0: 4})

    #
    # VAL TESTING
    #
    # index -> (lambda row: row[index])
    def test_index_val_replace(self):
        cmd = xlmp.mpCmd(i_INDEX_COMMAND, int_is_index=True)
        self.assertEqual(
            [cmd[k](TEST_ROW) for k in cmd.keys()],
            [o_INDEX_COMMAND[k](TEST_ROW) for k in o_INDEX_COMMAND.keys()]
            )

    # name -> (lambda row: row[index])
    def test_name_val_replace(self):
        cmd = xlmp.mpCmd(i_NAME_VAL_COMMAND, str_is_name=True)
        self.assertEqual(
            [cmd[k](TEST_ROW) for k in cmd.keys()],
            [o_INDEX_COMMAND[k](TEST_ROW) for k in o_INDEX_COMMAND.keys()]
            )

    # val -> (lambda *args, **kwargs: val)
    def test_value_val_replace(self):
        cmd = xlmp.mpCmd({0: TEST_STRING, 1: 1},
                         str_is_name=False, int_is_index=False)
        self.assertEqual(cmd[0](TEST_ROW), TEST_STRING)
        self.assertEqual(cmd[1](TEST_ROW), 1)

    # (func, (indexes)) -> (lambda row: func(*rmap((lambda i: row[i]), indexes)
    def test_func_val_replace(self):
        cmd = xlmp.mpCmd({
            0: (sum, ([0, 2, 4], 0)),
            1: (''.join, [PARROT_INDEXES])
            }, int_is_index=True)
        self.assertEqual(cmd[0](TEST_INT_ROW), 9001)
        self.assertEqual(cmd[1](string.ascii_uppercase), TEST_STRING)

    # (func, (offset_indexes)) ->
    #    (lambda row: func(*rmap((lambda i: row[i], indexes)
    def test_func_val_offset(self):
        cmd = xlmp.mpCmd({
            'A': (sum, ([1, 3, 5], 1)),
            'B': (''.join, [[i + 1 for i in PARROT_INDEXES]])
            }, int_is_index=True, offset=1)
        self.assertEqual(cmd[0](TEST_INT_ROW), 9001)
        self.assertEqual(cmd[1](string.ascii_uppercase), TEST_STRING)

    # (func, (names)) -> (lambda row: func(*rmap((lambda i: row[i], indexes)
    def test_func_val_names(self):
        cmd = xlmp.mpCmd({
            0: (sum, (['A', 'C', 'E'], 'A')),
            1: (''.join, [[letter for letter in TEST_STRING]])
            }, int_is_index=True)
        self.assertEqual(cmd[0](TEST_INT_ROW), 9001)
        self.assertEqual(cmd[1](string.ascii_uppercase), TEST_STRING)

    # entry point testing
    #  __ini__ already tested
    #  __setitem__
    def test_setitem(self):
        cmd = xlmp.mpCmd({}, int_is_index=True, offset=1)
        cmd[1] = 1
        cmd['B'] = 'A'
        self.assertEqual(cmd[0](TEST_ROW), TEST_ROW[0])

    # dict.update
    def test_update(self):
        other_dict = {1: 1, 'B': 'A'}
        cmd = xlmp.mpCmd({}, offset=1, int_is_index=True, str_is_name=True)
        cmd.update(other_dict)
        self.assertEqual(cmd[0](TEST_ROW), TEST_ROW[0])
        self.assertEqual(cmd[1](TEST_ROW), TEST_ROW[0])


if __name__ == '__main__':
    unittest.main()
