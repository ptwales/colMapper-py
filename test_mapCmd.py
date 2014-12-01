from xlmp import mapCmd
import string
import unittest


class TestFloatingFuncs(unittest.TestCase):

    def test_rmap(self):
        self.assertEqual(mapCmd.rmap(lambda x: x, []), [])
        x = mapCmd.rmap((lambda x: 1 + x),  [1, (2, 3), [4, (5, 6)]])
        self.assertEqual(x, [2, [3, 4], [5, [6, 7]]])
    
    def test_name_to_index(self):
        self.assertEqual(mapCmd.name_to_index('A'), 0)
        self.assertEqual(mapCmd.name_to_index('AA'), 26)
        self.assertEqual(mapCmd.name_to_index('XFD'), 16383)
        self.assertRaises(TypeError, mapCmd.name_to_index, '')


class TestMapCmdKey(unittest.TestCase):
    
    # stupid test case; nothing should change
    # index -> index
    def test_index_key(self):
        cmd = mapCmd.mapCmd({0: 0})
        self.assertEqual(cmd[0](['a']), 'a')

    # index -> index - offset
    def test_index_key_offset(self):
        cmd = mapCmd.mapCmd({1: 1}, offset=1)
        self.assertEqual(cmd[0](['a']), 'a')

    # name -> index
    def test_name_key_replace(self):
        cmd = mapCmd.mapCmd({'A': 0, 1: 'B'}, str_is_name=True)
        self.assertEqual(cmd[0](['a', 'b']), 'a')
        self.assertEqual(cmd[1](['a', 'b']), 'b')
        
    # there is no name offset
    # name -> index
    def test_name_key_offset(self):
        cmd = mapCmd.mapCmd({'A': 1, 2: 'B'}, str_is_name=True, offset=1)
        self.assertEqual(cmd[0](['a', 'b']), 'a')
        self.assertEqual(cmd[1](['a', 'b']), 'b')

    # keys can only be int or str
    def test_invalid_key(self):
        with self.assertRaises(TypeError):
            mapCmd.mapCmd({(1, 2): 0}, str_is_name=True, int_is_index=True)
        
        
        
class TestMapCmdValue(unittest.TestCase):
    
    # index -> (lambda row: row[index])
    def test_index_val_replace(self):
        pass

    # name -> (lambda row: row[index])
    def test_name_val_replace(self):
        pass

    # val -> (lambda *args, **kwargs: val)
    def test_value_val_replace(self):
        pass

    # (func, (indexes)) -> (lambda row: func(*rmap((lambda i: row[i]), indexes)
    def test_func_val_replace(self):
        pass

    # (func, (offset_indexes)) ->
    #    (lambda row: func(*rmap((lambda i: row[i], indexes)
    def test_func_val_offset(self):
        pass

    # (func, (names)) -> (lambda row: func(*rmap((lambda i: row[i], indexes)
    def test_func_val_names(self):
        pass



class TestMapCmdEntryPoints(unittest.TestCase):
    
    #  __init__ already tested
    #  __setitem__
    def test_setitem(self):
        pass
      
    # dict.update
    def test_update(self):
        pass




if __name__ == '__main__':
    unittest.main()