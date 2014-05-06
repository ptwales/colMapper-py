import unittest
import xlmp


TEST_MATRIX = [[unicode(string.uppercase[c] + str(r + 1), 'utf8')
                for c in DEFAULT_RANGE] for r in DEFAULT_RANGE]
DEFAULT_RANGE = range(0, 26)
INDEX_COMMAND = {i: i for i in DEFAULT_RANGE}
NAME_COMMAND = {string.uppercase[c]: i for i in DEFAULT_RANGE}


class TestmpCmd(unittest.TestCase):

    def test_index_offset(self):
        for offset_ in DEFAULT_RANGE:
            test_cmd = xlmp.mpCmd(INDEX_COMMAND, offset=offset_)    
            self.assertEqual(test_cmd, 
                {i + offset_ : i + offset_ for i in DEFAULT_RANGE}
                )
                
    def test_name_replace(self):
        test_cmd = xlmp.mpCmd(NAME_COMMAND)    
        self.assertEqual(test_cmd, INDEX_COMMAND)
    
    def test_name_offset(self):
        for offset_ in DEFAULT_RANGE:
            test_cmd = xlmp.mpCmd(NAME_COMMAND, offset=offset_)    
            self.assertEqual(test_cmd, 
                {i : i + offset_ for i in DEFAULT_RANGE}
                )
                
        
if __name__ == '__main__':
    unittest.main()
