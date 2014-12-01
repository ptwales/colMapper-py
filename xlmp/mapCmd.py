from functools import reduce

def rmap(func, sequence):
    return [rmap(func, i) if isinstance(i, (tuple, list))
            #elif isinstance(i, dict) ???
            else func(i)
            for i in sequence]


def name_to_index(col_name):
    # name_to_index('') = -1. xlwt will throw an error but that will
    # occur too far down the line.  Catch it now.
    if not col_name:
        raise TypeError("Null column names are not allowed")
    # if len(col_name) > 3:
    #   raise TypeError("{} is too long to be a column name".format(col_name))
    return reduce((lambda index, char: index * 26 + int(char, 36) - 9),
                  col_name, 0) - 1


class mapCmd(dict):
    """Stores user defined maps and converts them to f(vector) = scalar

    Every item stored in mapCmd will be converted to int: (lambda row: some_func)
    """

    def __init__(self, map_dict, offset=0,
            int_is_index=True, str_is_name=False):
        self.offset = offset
        self.int_is_index = int_is_index
        self.str_is_name = str_is_name
        super(mapCmd, self).__init__(self._convert_dict(map_dict))

    # Override setters, do no override accessors
    def __setitem__(self, key, val):
        super(mapCmd, self).__setitem__(*self._convert_item(key, val))

    def update(self, other):
        super(mapCmd, self).update(self._convert_dict(other))

    # Macro functions
    def _convert_dict(self, other):
        return {self._convert_key(key): self._convert_val(val)
                    for key, val in other.items()}

    def _convert_item(self, key, val):
        return self._convert_key(key), self._convert_val(val)

    # actual replacement
    def _convert_val(self, val):

        if ((isinstance(val, int) and self.int_is_index) or
                (isinstance(val, str) and self.str_is_name)):
            return (lambda row, index=self._convert_key(val): row[index])

        elif isinstance(val, (tuple, list)):
            func, indexes = val
            if not callable(func):
                raise TypeError("{} is not callable.".format(func))
            indexes = rmap(self._convert_key, indexes)
            return (lambda row: func(*rmap((lambda i: row[i]), indexes)))

        else:
            return (lambda *args, **kwargs: val)

    def _convert_key(self, key):
        if isinstance(key, str):
            return name_to_index(key)
        elif isinstance(key, int):
            return key - self.offset
        else:
            raise TypeError("""\
            {} is an invalid index; it is not type str or int.\
            """.format(key))


