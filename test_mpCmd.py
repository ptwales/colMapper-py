import unittest
import xlmp
import string


DEFAULT_LIMIT = 26
DEFAULT_RANGE = range(0, DEFAULT_LIMIT)
TEST_MATRIX = [[string.uppercase[c] + str(r + 1)
                for c in DEFAULT_RANGE] for r in DEFAULT_RANGE]

i_INDEX_COMMAND = {i: i for i in DEFAULT_RANGE}
i_NAME_COMMAND = {string.uppercase[i]: i for i in DEFAULT_RANGE}
i_OFFSET_COMMAND = {i + 1: i + 1 for i in DEFAULT_RANGE}

o_INDEX_COMMAND = {i: (lambda row, i=i: row[i]) for i in DEFAULT_RANGE}

TEST_ROW = TEST_MATRIX[15]
TEST_STRING = "This is an ex-parrot"


class TestmpCmd(unittest.TestCase):

    # stupid test case; nothing should change
    def test_index_key(self):
        test_cmd = xlmp.mpCmd(i_INDEX_COMMAND)
        self.assertEqual(test_cmd.keys(), o_INDEX_COMMAND.keys())

    def test_index_key_offset(self):
        test_cmd = xlmp.mpCmd(i_OFFSET_COMMAND, off_set=1)
        self.assertEqual(test_cmd.keys(), o_INDEX_COMMAND.keys())

    def test_name_replace(self):
        test_cmd = xlmp.mpCmd(i_NAME_COMMAND)
        self.assertEqual(test_cmd.keys(), o_INDEX_COMMAND.keys())

    def test_name_offset(self):
        test_cmd = xlmp.mpCmd(i_NAME_COMMAND, off_set=1)
        self.assertEqual(test_cmd.keys(), o_INDEX_COMMAND.keys())

    def test_invalid_key(self):
        with self.assertRaises(TypeError):
            test_cmd = xlmp.mpCmd({[1, 2, 3]: 4})

    def test_index_val_replace(self):
        test_cmd = xlmp.mpCmd(i_INDEX_COMMAND)
        self.assertEqual(
            [test_cmd[k](TEST_ROW) for k in test_cmd.keys()],
            [o_INDEX_COMMAND[k](TEST_ROW) for k in o_INDEX_COMMAND.keys()]
            )

    def test_value_val_replace(self):
        test_cmd = xlmp.mpCmd({0: TEST_STRING})
        self.assertEqual(test_cmd[0](TEST_ROW), TEST_STRING)

if __name__ == '__main__':
    unittest.main()
