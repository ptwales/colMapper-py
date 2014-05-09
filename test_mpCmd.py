import unittest
import xlmp


DEFAULT_LIMIT = 26
DEFAULT_RANGE = range(0, DEFAULT_LIMIT)
TEST_MATRIX = [[unicode(string.uppercase[c] + str(r + 1), 'utf8')
                for c in DEFAULT_RANGE] for r in DEFAULT_RANGE]
INDEX_COMMAND = {i: i for i in DEFAULT_RANGE}
NAME_COMMAND = {string.uppercase[c]: i for i in DEFAULT_RANGE}
TEST_ROW = TEST_MATRIX[15]
TEST_STRING = "This is an ex-parrot"


class TestmpCmd(unittest.TestCase):

    def test_index_offset(self):
        for offset_ in DEFAULT_RANGE:
            test_cmd = xlmp.mpCmd(INDEX_COMMAND, offset=offset_)    
            self.assertEqual(
                test_cmd, 
                {i + offset_ : i for i in DEFAULT_RANGE}
                )
                
    def test_name_replace(self):
        test_cmd = xlmp.mpCmd(NAME_COMMAND)    
        self.assertEqual(test_cmd, INDEX_COMMAND)
    
    def test_name_offset(self):
        for offset_ in DEFAULT_RANGE:
            test_cmd = xlmp.mpCmd(NAME_COMMAND, offset=offset_)    
            self.assertEqual(
                test_cmd, 
                INDEX_COMMAND
                )
                
    def test_evaluate_index(self):
        for offset_ in DEFAULT_RANGE:
            for index in range(offset_, DEFAULT_LIMIT):
                self.assertEqual(
                    xlmp.mpCmd({}, offset_).operate.evaluate(index, TEST_ROW),
                    TEST_ROW[index + offset_]
                    )
    
    def test_evaluate_value(self):
        self.assertEqual(
            xlmp.mpCmd({}).operate.evaluate(TEST_STRING, TEST_ROW),
            TEST_STRING
            )

    def test_evaluate_func(self):
        
        def finite_args(arg1, arg2, arg3):
            return arg1 + arg2 + arg3
            
        blank_command = xlmp.mpCmd({})
        simple_args = [1, 2, 3]
        for func in [finite_args, sum]:
            self.assertEqual(
                blank_command.operate.evaluate((func, simple_args)),
                6
                )
                
        
if __name__ == '__main__':
    unittest.main()
