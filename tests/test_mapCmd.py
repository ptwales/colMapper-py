from xlmp import mapCmd
# import string
import unittest


class TestFloatingFuncs(unittest.TestCase):

    def test_rmap(self):
        self.assertEqual(mapCmd.rmap(lambda x: x, []), [])
        result = mapCmd.rmap((lambda x: 1 + x),  [1, (2, 3), [4, (5, 6)]])
        self.assertEqual(result, [2, [3, 4], [5, [6, 7]]])
    
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
        self.assertEqual([0], list(cmd))

    # index -> index - offset
    def test_index_key_offset(self):
        cmd = mapCmd.mapCmd({1: 1}, offset=1)
        self.assertEqual([0], list(cmd))

    # name -> index
    def test_name_key_replace(self):
        cmd = mapCmd.mapCmd({'A': 0, 1: 'B'}, str_is_name=True)
        self.assertEqual([0, 1], list(cmd))

    # there is no name offset
    # name -> index
    def test_name_key_offset(self):
        cmd = mapCmd.mapCmd({'A': 1, 2: 'B'}, 
                            str_is_name=True,
                            offset=1)
        self.assertEqual([0, 1], list(cmd))
        

    # keys can only be int or str
    def test_invalid_key(self):
        with self.assertRaises(TypeError):
            mapCmd.mapCmd({(1, 2): 0}, str_is_name=True, int_is_index=True)


class TestMapCmdValue(unittest.TestCase):

    # index -> (lambda row: row[index])
    def test_index_val_replace(self):
        cmd = mapCmd.mapCmd({0: 0})
        self.assertEqual(cmd[0](['x']), 'x')

    # name -> (lambda row: row[index])
    def test_name_val_replace(self):
        cmd = mapCmd.mapCmd({'a': 0})
        self.assertEqual(cmd[0](['x']), 'x')

    # val -> (lambda *args, **kwargs: val)
    def test_value_val_replace(self):
        cmd = mapCmd.mapCmd({0: 'spam'})
        self.assertEqual(cmd[0](['x']), 'spam')

    # (func, (indexes)) -> (lambda row: func(*rmap((lambda i: row[i]), indexes)
    def test_func_val_replace(self):

        argless = mapCmd.mapCmd({0: (lambda: 'spam', [])}, offset=0)
        self.assertEqual(argless[0](['x']), 'spam')

        arged = mapCmd.mapCmd({0: ((lambda x: 'spam'), [0])})
        self.assertEqual(arged[0](['x']), 'spam')

        cmd = mapCmd.mapCmd({0: ((lambda x: x + 2), [0])})
        self.assertEqual(cmd[0]([2]), 4)

    # (func, (offset_indexes)) ->
    #    (lambda row: func(*rmap((lambda i: row[i], indexes)
    def test_func_val_offset(self): 
        cmd = mapCmd.mapCmd({1: ((lambda x: x + 2), [2])}, offset=1)
        self.assertEqual(cmd[0]([1,2,3,4]), 4)

    # (func, (names)) -> (lambda row: func(*rmap((lambda i: row[i], indexes)
    def test_func_val_names(self):
        cmd = mapCmd.mapCmd({'A': ((lambda x: x + 2), ['B'])}, str_is_name=True)
        self.assertEqual(cmd[0]([1,2,3,4]), 4)



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
