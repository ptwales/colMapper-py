import xlrd
import xlwt
from itertools import groupby

__all__ = ['xlmp', 'mpCmd', 'group_by_ids', 'Ixl']  #, 'xlSmp']


class Ixl(object):
    """xlmp's interface to the xlrd and xlwt modules.
    
    Ixl reads a rectangular list from a sheet or write one
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
        Initilizes the Ixl object with boolean by_row of whether sub-lists
        should be rows or columns of read sheets.
        
        Args:
            read_by_row: If true sub-lists are rows else columns
        """
        self.by_row = read_by_row

    def __transpose(self, M):
        """Transposes M if reading by column.
        
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
            return M
        else:
            return [list(a) for a in zip(*M)]

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
        M = [sheet.row_values(r, c0, cf) for r in range(r0, rf)]
        return self.__transpose(M)

    def read_book(self, book_path, sheet_i=0, **kw_ranges):
        """Opens book_path and reads dataset from sheet_i.
        
        Reads the file at book_path and gets sheet_i of it.
        sheet_i may be a index or a name of the sheet in the book to read.
        
        Args:
            book_path: The path to the workbook file to read.
            sheet_i: The zero-offset index or name of the sheet
            of the book to read.
            c0=0, cf=-1, r0=0, rf=-1: additional arguments to pass to read_sheet.
        
        Returns: 
            Rectangular list of cell values of the sheet_i sheet of
            the book at book_path within the ranges passed in c0=0, cf=-1, r0=0, rf=-1.
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
            c0=0, cf=-1, r0=0, rf=-1: The additional arguements passed to either 
            read_sheet or read_book.
        """
        try:
            M = self.read_sheet(sheet, **kw_ranges)
        except AttributeError:
            M = self.read_book(book_path, sheet, **kw_ranges)
        return M

    def write_sheet(self, M, sheet, c0=0, r0=0):
        """Writes dataset M to sheet starting at (c0, r0)
        
        Writes dataset M to xlwt sheet object starting at row=r0
        and column=c0 and extending to the left and down, (increasing
        row and column indexes). If set to read by_row then sub lists
        of M will be rows otherwise sublists of M will be columns.
        
        Args:
            M: dataset to write. A list of lists is expected.
            sheet: xlrwt.Sheet object where the dataset M will be written.
            c0: Initial column index of sheet to write the dataset.
            r0: Initial row index of sheet to write the dataset.
        """
        M = self.__transpose(M)
        for r, row, in enumerate(M):
            for c, el in enumerate(row):
                sheet.write(r0 + r, c0 + c, el)

    def write_book(self, M, book_path, sheet_name='xlmp',  **kw_to_point):
        """Writes dataset M to new workbook file at book_path.
        
        
        """
        book = xlwt.Workbook()
        sheet = book.add_sheet(sheet_name)
        self.write_sheet(M, sheet, **kw_to_point)
        book.save(book_path)

    def guess_write(self, M, sheet, book_path='', **kw_to_point):
        """Writes dataset M to existing sheet or new workbook file at book_path.
        """
        try:
            self.write_sheet(M, sheet,  c0, r0)
        except AttributeError:
            self.write_book(M, book_path, sheet, **kw_to_point)


class mpCmd(dict):
    """Stores user defined maps and performs the map operations
    """
    
    def __init__(self, map_dict, offset=0):
        self._Off = offset
        # just call __setitem__ and be done with it
        for k in map_dict.keys():
            self[k] = map_dict[k]
    
    def __setitem__(self, key, val):
        if isinstance(key, str):
            key = self.__replace_col_name(key)
        elif isinstance(key, int):
            key += self._Off  # names are not offset
        else:
            raise TypeError
        super(mpCmd, self).__setitem__(key, val)
    
# Still in ToX style documentation
    # Converts all of the keys that are strings to
    # integers, assuming that strings are xl colnames.
    # unnecessary function only for client side conviencence.
    def __replace_col_name(self, key):
        place = 1
        index = 0
        for char in reversed(key):
            index += place * (int(char.upper(), 36) - 9)
            place *= 26
        index -= 1
        return index
    
    # performs [M].[F] = B
    # where F_j(M_i) = B[i][j]
    # default is by cols of F and rows of M
    # M must be transposed beforehand for other method
    def operate(self, M):
        # Source of my woes
        # returns $f(\vec{r})$
        # unless f is a string then it returns f
        # or if f is an int then it returns r[f]
        # Need a better structure for handeling functions
        #
        # TODO: specialize indexes instead of strings and ints
        # new class def may be needed.
        def evaluate(self, var, row):
            if isinstance(var, str):
                return var
            elif isinstance(var, int):
                return row[var - self._Off]
            else:  # this looks like horrible recursion
                return var[0](*[evaluate(v, row) for v in var[1]])
            
        x_range = range(max(self.keys()) + 1)
        y_range = range(len(M[0]))
        B = [[None for i in x_range] for j in y_range]
        assert len(x_range) == len(M)
        for i, row in enumerate(M):
            for k in self.keys():
                B[i][k] = self.evaluate(self[k], row)
        return B


def xlmp(cmd, map_by_col=True, f_book='', t_book='', f_sheet=0, t_sheet='xlmp',
         fc0=0,  fcf=-1, fr0=0, frf=-1, tr0=0, tc0=0):
             
    xlrw = Ixl(read_by_row=map_by_col)
    '''
    M = xlrw.guess_read(f_book, f_sheet, fc0, fcf, fr0, frf)
    B = cmd.operate(M)
    xlrw.guess_write(B, t_book, t_sheet, tr0, tc0)
    '''
    xlrw.guess_write(cmd.operate(xlrw.guess_read(f_book, f_sheet, fc0, fcf, fr0, frf)),
                    t_book, t_sheet, tr0, tc0)



def group_by_ids(M, id_indexes):
    id_func = lambda x: [x[i] for i in id_indexes]
    M.sort(key=id_indexes)
    return [list(g) for k, g in groupby(M, id_func)]


# Still mock up
def xlsmp(sub_cmd, grp_func, grp_by_col=True, f_book='', t_book='', 
          f_sheet=0, t_sheet='xlmp', fc0=0, fcf=-1, fr0=0, frf=-1,
          tr0=0, tc0=0, **kw_grpargs):
              
    xlrw = Ixl(read_by_row=grp_by_col)
    M = xlrw.guess_read(f_book, f_sheet, fc0, fcf, fr0, frf)
    # must we map along the same dim that as grouped?
    B = zip([sub_cmd.operate(zip(*m)) for m in grp_func(M, **kw_grpargs)])
    xlrw.guess_write(B, t_book, t_sheet, tr0, tc0)
