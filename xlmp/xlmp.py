import xlrd
import xlwt
from itertools import groupby

__all__ = ['xlmp', 'mpCmd', 'groupByIds']  #, 'xlSmp']

##
# XLRD XLWT Interfaces
class ixl(object):
    """xlmp's interface to the xlrd and xlwt modules."""
    def __init__(self, readByRow=False):
        self.byRow = readByRow


    def __transpose(self, M):
        return M if self.byRow else list(zip(*M))


    ##
    # creates matrix of values of xlrd.Sheet object
    def pullSheet(self, s, cRng=(0, -1), rRng=(0, -1)):
        assert not s.ragged_rows
        rRng = list(rRng)
        if rRng[1] < 0:
            rRng[1] += s.nrows
        M = [s.row_values(r, *cRng) for r in range(*rRng)]
        return self.__transpose(M)


    def readBook(self, book_path, sheet_i=0, *args):
        book = xlrd.open_workbook(book_path)
        try:
            sheet = book.sheet_by_index(sheet_i)
        except TypeError:
            sheet = book.sheet_by_name(sheet_i)
        return self.pullSheet(sheet, *args)


    def guessRead(self, book, sheet, *args):
        try:
            M = self.readSheet(sheet, *args)
        except AttributeError:
            M = self.readBook(book, sheet, *args)
        return M

    ##
    # writes maxtrix M to xlwt.Sheet object
    # offset by r0 and c0
    def writeSheet(self, M, s, c0=0, r0=0):
        M = self.__transpose(M)
        for r in range(len(M)):
            for c in range(len(M[r])):
                s.write(r0 + r, c0 + c, M[r][c])


    def writeBook(self, M, book_path, sheet_name='xlmp', c0=0, r0=0):
        book = xlwt.Workbook()
        sheet = book.add_sheet(sheet_name)
        self.writeSheet(M, sheet, c0, r0)
        book.save(book_path)


    def guessWrite(self, M, book, sheet, *args):
        try:
            self.writeSheet(M, sheet, *args)
        except AttributeError:
            self.writeBook(M, book, sheet, *args)


class mpCmd(dict):


    def __init__(self, cmd, offset=0):
        self._Off = offset
        for k in cmd.keys():  # call __setitem__ and don't redefine
            self[k] = cmd[k]


    def __setitem__(self, key, val):
        if isinstance(key, str):
            key = self.__ReplaceColName(key)
        elif isinstance(key, int):
            key += self._Off
        else:
            raise TypeError
        super(mpCmd, self).__setitem__(key, val)


    ##
    # Converts all of the keys that are strings to
    # integers, assuming that strings are xl colnames.
    # unnecessary function only for client side conviencence.
    def __ReplaceColName(self, key):
        place = 1
        index = 0
        for c in reversed(key):
            index += place * (int(c.upper(), 36) - 9)
            place *= 26
        index -= 1
        return index
        
        
    ##
    # Source of my woes
    # returns $f(\vec{r})$
    # unless f is a string then it returns f
    # or if f is an int then it returns r[f]
    # Need a better structure for handeling functions
    def evaluate(self, var, row):
        if isinstance(var, str):
            return var
        elif isinstance(var, int):
            return row[var - self._Off]
        else:  # this looks like horrible recursion
            return var[0](*[evaluate(v, row) for v in var[1]])
            

    ##
    # performs [M].[F] = B
    # where F_j(M_i) = B[i][j]
    # default is by cols of F and rows of M
    # M must be transposed beforehand for other method
    def operate(self, M):
        X = list(range(max(self.keys()) + 1))
        Y = list(range(len(M[0])))
        B = [[None for i in X] for j in Y]
        for i, r in zip(X, M):
            for k in self.keys():
                B[i][k] = self.evaluate(self[k], r)
        return B



def xlmp(cmd, mapByCol=True, fBook='', tBook='', fSheet=0, tSheet='xlmp',
         fCrng=(0, -1), fRrng=(0, -1), tp=(0, 0):
    xlrw = ixl(readByRow=mapByCol)
    # these could just be stacked in one big overload
    M = xlrw.guessRead(fBook, fSheet, fCrng, fRrng)
    B = cmd.operate(M)
    xlrw.guessWrite(B, tBook, tSheet, *tp)
    '''
    xlrw.guessWrite(cmd.operate(xlrw.guessRead(fBook, fSheet, fCrng, fRrng)),
                    tBook, tSheet, *tp)
    '''


def groupByIds(M, idLocs):
    idFunc = lambda x: [x[d] for d in idLocs]
    M.sort(key=idFunc)
    return [list(g) for k, g in groupby(M, idFunc)]


# Still mock up
def xlSmp(subCmd, grpFnc, grpByCol=True, fBook='', tBook='', fSheet=0, tSheet='xlmp',
          fCrng=(0, -1), fRrng=(0, -1), tp=(0, 0), **grpFncArgs):
    xlrw = ixl(readbyRow=grpByCol)
    M = xlrw.guessRead(fBook, fSheet, fCrng, fRrng)
    gM = grpFnc(M, grpFncArgs)  # I don't know if this is how to use **kwargs
    gB = [subCmd.operate(zip(*m)) for m in gM]  # must map along the same dim that as grouped?
    B = zip(gB)
    xlrw.guessWrite(B, tBook, tSheet, *tp)
    '''
    xlrw.guessWrite(zip(
                    [subCmd.operate(zip(*m)) for m in 
                    grpFng(xlrw.guessRead(fBook, fSheet, fCrng, fRrng))]
                    ), tBook, tSheetm *tp)
    '''
