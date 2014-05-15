import unittest
import xlmp
import string


DEFAULT_LIMIT = 26
DEFAULT_RANGE = range(0, DEFAULT_LIMIT)
TEST_MATRIX = [[string.ascii_uppercase[c] + str(r + 1)
                for c in DEFAULT_RANGE] for r in DEFAULT_RANGE]

i_INDEX_COMMAND = {i: i for i in DEFAULT_RANGE}
i_NAME_COMMAND = {string.ascii_uppercase[i]: i for i in DEFAULT_RANGE}
i_OFFSET_COMMAND = {i + 1: i + 1 for i in DEFAULT_RANGE}

o_INDEX_COMMAND = {i: (lambda row, i=i: row[i]) for i in DEFAULT_RANGE}

TEST_ROW = TEST_MATRIX[15]
TEST_STRING = "This is an ex-parrot"


class TestmpCmd(unittest.TestCase):

    #
    # KEY TESTING
    #
    # stupid test case; nothing should change
    # index -> index
    def index_key(self):
        cmd = xlmp.mpCmd(i_INDEX_COMMAND)
        self.assertEqual(cmd.keys(), o_INDEX_COMMAND.keys())

    # index -> index - offset
    def index_key_offset(self):
        cmd = xlmp.mpCmd(i_OFFSET_COMMAND, offset=1)
        self.assertEqual(cmd.keys(), o_INDEX_COMMAND.keys())

    # name -> index
    def name_key_replace(self):
        cmd = xlmp.mpCmd(i_NAME_COMMAND)
        self.assertEqual(cmd.keys(), o_INDEX_COMMAND.keys())

    # there is no name offset
    # name -> index
    def name_key_offset(self):
        cmd = xlmp.mpCmd(i_NAME_COMMAND, offset=1)
        self.assertEqual(cmd.keys(), o_INDEX_COMMAND.keys())

    # keys can only be int or str
    def invalid_key(self):
        with self.assertRaises(TypeError):
            cmd = xlmp.mpCmd({[1, 2, 3]: 4})
            cmd = xlmp.mpCmd({2.0: 4})

    # 
    # VAL TESTING
    #
    # index -> (lambda row: row[index])
    #
    def index_val_replace(self):
        cmd = xlmp.mpCmd(i_INDEX_COMMAND, int_is_name=True)
        self.assertEqual(
            [cmd[k](TEST_ROW) for k in cmd.keys()],
            [o_INDEX_COMMAND[k](TEST_ROW) for k in o_INDEX_COMMAND.keys()]
            )

    # name -> (lambda row: row[index])
    def name_val_replace(self):
        cmd = xlmp.mpCmd(i_NAME_COMMAND, str_is_name=True)
        self.assertEqual(
            [cmd[k](TEST_ROW) for k in cmd.keys()],
            [o_INDEX_COMMAND[k](TEST_ROW) for k in o_INDEX_COMMAND.keys()]
            )

    # val -> (lambda *args, **kwargs: val)
    def value_val_replace(self):
        cmd = xlmp.mpCmd({0: TEST_STRING, 1: 1},
                         str_is_name=False, int_is_index=False)
        self.assertEqual(cmd[0](TEST_ROW), TEST_STRING)
        self.assertEqual(cmd[1](TEST_ROW), 1)


    # (func, (indexes)) -> (lambda row: func(*rmap((lambda i: row[i]), indexes)
    def func_val_replace(self):
        cmd = xlmp.mpCmd({0: (sum, ([0, 2, 4], 0))}, int_is_index=True) 
        self.assertEqual(cmd[0]([1, -1600, 5000, 4999]), 9001)
        #cmd = xlmp.mpCmd({0, (''.join, ([18))


if __name__ == '__main__':
    unittest.main()
