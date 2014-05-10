import unittest
import xlmp
import string


DEFAULT_LIMIT = 26
DEFAULT_RANGE = range(0, DEFAULT_LIMIT)
TEST_MATRIX = [[string.uppercase[c] + str(r + 1)
                for c in DEFAULT_RANGE] for r in DEFAULT_RANGE]
INDEX_COMMAND = {i: i for i in DEFAULT_RANGE}
NAME_COMMAND = {string.uppercase[i]: i for i in DEFAULT_RANGE}
OFFSET_COMMAND = {i + 1: i for i in DEFAULT_RANGE}
TEST_ROW = TEST_MATRIX[15]
TEST_STRING = "This is an ex-parrot"



class TestmpCmd(unittest.TestCase):

    def test_index_offset(self):
        test_cmd = xlmp.mpCmd(OFFSET_COMMAND, offset=1) 
        self.assertEqual(test_cmd, INDEX_COMMAND)
                
    def test_name_replace(self):
        test_cmd = xlmp.mpCmd(NAME_COMMAND)    
        self.assertEqual(test_cmd, INDEX_COMMAND)
    
    def test_name_offset(self):
        test_cmd = xlmp.mpCmd(NAME_COMMAND, offset=1)    
        self.assertEqual(test_cmd, INDEX_COMMAND)
            
    def test_evaluate_index(self):
        for offset_ in DEFAULT_RANGE:
            for index in range(offset_, DEFAULT_LIMIT):
                self.assertEqual(
                    xlmp.mpCmd({}, offset_).evaluate(index, TEST_ROW),
                    TEST_ROW[index - offset_]
                    )
    
    def test_evaluate_value(self):
        self.assertEqual(
            xlmp.mpCmd({}).evaluate(TEST_STRING, TEST_ROW),
            TEST_STRING
            )


#    def test_evaluate_func(self):
#        
#        def finite_args(arg1, arg2, arg3):
#            return arg1 + arg2 + arg3
#            
#        blank_command = xlmp.mpCmd({})
#        for func in [finite_args, sum]:
#            self.assertEqual(
#                blank_command.evaluate((func, [0, 1, 2]), 
#                                       [1, 2, 3]),
#                6
#                )

        
if __name__ == '__main__':
    unittest.main()
