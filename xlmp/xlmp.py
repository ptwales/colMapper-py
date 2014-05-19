import xlrd
import xlwt
from itertools import groupby
import collections

__all__ = ['xlmp', 'mpCmd', 'group_by_ids', '_ExcelInterface']  # , 'xlSmp']


class _ExcelInterface(object):
    """xlmp's interface to the xlrd and xlwt modules.

    ExcelInterface reads a rectangular list from a sheet or write one
    to the sheet.  It also transposes the list in reading
    and writing if by_row = False.
    Sheets read must not have ragged rows.

    Attributes:
        by_row: Boolean whether sub-lists are rows (T) or columns (F).
        Checked at reading and writing of datasets.  If False, datasets
        are transposed when read and written.  This way xlmp and mpCmd
        can ignore transposition.
    """

    def __init__(self, read_by_row=True):
        """
        Initilizes the ExcelInterface object with boolean by_row
        of whether sub-lists should be rows or columns of read sheets.

        Args:
            read_by_row: If true sub-lists are rows else columns
        """
        self.by_row = read_by_row

    def __transpose(self, data_matrix):
        """Transposes data_matrix if reading by column.

        Private method to transpose datasets, if sub-lists
        correspond to columns and not rows, before reading
        or writing.

        Args:
            M: data set to transpose.  Should be a rectangular
            2D list.

        Returns:
            M if if reading by row else it return the transpose of M
        """
        if self.by_row:
            return data_matrix
        else:
            return [list(a) for a in zip(*data_matrix)]

    def read_sheet(self, sheet, c0=0, cf=-1, r0=0, rf=-1):
        """Reads a rectangular list from sheet.

        Reads the cell values of sheet in column range [c0, cf]
        and row range [r0, rf] and returns the dataset as as a
        rectangular list.

        Args:
            sheet: The xlrd.Sheet object that is read.
            c0, cf: inclusive column range to read from sheet.
            r0, rf: the inclusive row range to read from sheet.

        Returns:
            A rectangular list of the cell values in the range
            ([c0, cf], [r0, rf]) of sheet.  Sub lists will be the rows
            if self.by_row is true, else they will be the columns.
        """
        # TODO: Raise an error instead. Research which error to raise
        assert not sheet.ragged_rows
        if rf < 0:
            rf += sheet.nrows + 1
        # xlrd.sheet.row_values uses row[c0:cf]
        # therefore cf=-1 excludes last.
        if cf < 0:
            cf += sheet.ncols + 1
        data_matrix = [sheet.row_values(r, c0, cf) for r in range(r0, rf)]
        return self.__transpose(data_matrix)

    def read_book(self, book_path, sheet_i=0, **kw_ranges):
        """Opens book_path and reads dataset from sheet_i.

        Reads the file at book_path and gets sheet_i of it.
        sheet_i may be a index or a name of the sheet in the book to read.

        Args:
            book_path: The path to the workbook file to read.
            sheet_i: The zero-offset index or name of the sheet
            of the book to read.
            c0=0, cf=-1, r0=0, rf=-1: additional arguments to pass to
            read_sheet.
        Returns:
            Rectangular list of cell values of the sheet_i sheet of
            the book at book_path within the ranges passed in c0=0,
            cf=-1, r0=0, rf=-1.
        """
        book = xlrd.open_workbook(book_path)
        try:
            sheet = book.sheet_by_index(sheet_i)
        except TypeError:
            sheet = book.sheet_by_name(sheet_i)
        return self.read_sheet(sheet, **kw_ranges)

    def guess_read(self, sheet=0, book_path='', **kw_ranges):
        """Reads a dataset from sheet or book_path

        Trys to read sheet as an xlrd.Sheet object with
        read_sheet.  If sheet is not that object then it will
        read the book with read_book.  This function and guess_write
        should only be used in user interfaces so that users may
        open their own xlrd.Sheet objects and pass them quickly
        to xlmp.  Users cannot pass xlrd.book objects because they
        are expected to open the sheet objects.

        Args:
            sheet: Either xlrd.Sheet object to read or a sheet
            index or name of an sheet object in the book
            at book_path.
            book_path: Path to a workbook file. Optionally, ignored
            if sheet is an xlrd.Sheet object.
            c0=0, cf=-1, r0=0, rf=-1: The additional arguements passed
            to either read_sheet or read_book.
        """
        try:
            data_matrix = self.read_sheet(sheet, **kw_ranges)
        except AttributeError:
            data_matrix = self.read_book(book_path, sheet, **kw_ranges)
        return data_matrix

    def write_sheet(self, data_matrix, sheet, c0=0, r0=0):
        """Writes dataset data_matrix to sheet starting at (c0, r0)

        Writes dataset M to xlwt sheet object starting at row=r0
        and column=c0 and extending to the left and down, (increasing
        row and column indexes). If set to read by_row then sub lists
        of M will be rows otherwise sublists of M will be columns.

        Args:
            data_matrix: dataset to write. A list of lists is expected.
            sheet: xlrwt.Sheet object where the dataset M will be written.
            c0: Initial column index of sheet to write the dataset.
            r0: Initial row index of sheet to write the dataset.
        """
        data_matrix = self.__transpose(data_matrix)
        for r, row, in enumerate(data_matrix):
            for c, el in enumerate(row):
                sheet.write(r0 + r, c0 + c, el)

    def write_book(self, data_matrix, book_path, sheet_name='xlmp',
                  **kw_to_point):
        """Writes dataset data_matrix to new workbook file at book_path.

        """
        book = xlwt.Workbook()
        sheet = book.add_sheet(sheet_name)
        self.write_sheet(data_matrix, sheet, **kw_to_point)
        book.save(book_path)

    def guess_write(self, data_matrix, sheet, book_path='', **kw_to_point):
        """Writes dataset data_matrix to existing sheet or new workbook file at
        book_path.
        """
        try:
            self.write_sheet(data_matrix, sheet, **kw_to_point)
        except AttributeError:
            self.write_book(data_matrix, book_path, sheet, **kw_to_point)


def rmap(func, sequence):
    return [rmap(func, i) if isinstance(i, (tuple, list))
            #elif isinstance(i, dict) ???
            else func(i) 
            for i in sequence]

    
def name_to_index(col_name):
    return reduce((lambda index, char: index*26 + int(char, 36) - 9),
                  col_name, 0) - 1

    
class mpCmd(dict):
    """Stores user defined maps and converts them to f(vector) = scalar
    
    Every item stored in mpCmd will be converted to int: (lambda row: some_func)
    """

    def __init__(self, map_dict, offset=0, 
                 int_is_index=True, str_is_name=False):
        self.offset = offset
        self.int_is_index = int_is_index
        self.str_is_name = str_is_name
        super(mpCmd, self).__init__(self._convert_dict(map_dict))

    # Override setters, do no override accessors
    def __setitem__(self, key, val):
        super(mpCmd, self).__setitem__(*self._convert_item(key, val))

    def update(self, other):
        super(mpCmd, self).update(self._convert_dict(other))
                                 
    # Macro functions
    def _convert_dict(self, other):
        return {self._convert_key(key): self._convert_val(val)
                    for key, val in other.items()}    
                    
    def _convert_item(self, key, val):
        return self._convert_key(key), self._convert_val(val)

    # actual replacement
    def _convert_val(self, val):
        if (isinstance(val, int) and self.int_is_index
                or isinstance(val, str) and self.str_is_name):
            return (lambda row, index=self._convert_key(val): row[index])
            
        elif isinstance(val, (tuple, list)):
            func, indexes = val
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
            raise TypeError


# performs [M].[F] = B
# where F_j(M_i) = B[i][j]
# default is by cols of F and rows of M
# M must be transposed beforehand for other method
def operate(map_command, data_matrix):
    x_range = range(max(map_command.keys()) + 1)
    y_range = range(len(data_matrix[0]))
    # preallocate because some columns could be skipped
    resultant_matrix = [[None for i in x_range] for j in y_range]
    assert len(x_range) == len(data_matrix)
    for i, row in enumerate(data_matrix):
        for k in map_command.keys():
            resultant_matrix[i][k] = map_command[k](row)
    return resultant_matrix


# xlmp should not need to know the kwargs of guess_read and
# guess_write
def xlmp(cmd, map_by_col=True, f_book='', t_book='', f_sheet=0, t_sheet='xlmp',
         from_ranges={}, to_point={}):
    xl_interface = _ExcelInterface(read_by_row=map_by_col)
    data_matrix = xl_interface.guess_read(f_book, f_sheet, **from_ranges)
    mapped_matrix = operate(cmd, data_matrix)
    xl_interface.guess_write(mapped_matrix, t_book, t_sheet, **to_point)


def group_by_ids(data_matrix, id_indexes):

    def id_func(x):
        return [x[i] for i in id_indexes]

    data_matrix.sort(key=id_indexes)
    return [list(g) for k, g in groupby(data_matrix, id_func)]


# Still mock up
def xlsmp(sub_cmd, grp_func, grp_by_col=True, f_book='', t_book='',
          f_sheet=0, t_sheet='xlmp', from_ranges={}, to_point={},
          grp_func_kwargs={}):
    xl_interface = _ExcelInterface(read_by_row=grp_by_col)
    data_matrix = xl_interface.guess_read(f_book, f_sheet, **from_ranges)
    block_matrix = grp_func(data_matrix, **grp_func_kwargs)
    # map each block then merge the mapped_blocks
    mapped_matrix = zip([operate(cmd, zip(*block))
                         for block in block_matrix])
    xl_interface.guess_write(mapped_matrix, t_book, t_sheet, **to_point)
