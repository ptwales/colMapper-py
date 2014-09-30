import xlmp
import string
import unittest

DEFAULT_RANGE = range(0,26)

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

    #
    # ENTRY POINT TESTING
    #
    #  __init__ already tested
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
