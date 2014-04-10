import xlrd
import xlwt
from itertools import groupby


##
# XLRD XLWT Interfaces
class ixl(object):

    def __init__(self, readByRow=False):
        self.byRow = readByRow

    def __transpose(self, M):
        return M if self.byRow else list(zip(*M))

    ##
    # creates matrix of values of xlrd.Sheet object
    def pullSheet(self, s, cRng=(0, -1), rRng=(0, -1)):
        assert not s.ragged_rows
        rRng = list(*rRng)
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

    def __init__(self, zeroOffset, *args):
        self._zOff = 0 if zeroOffset else 1
        dict.__init__(*args)

    def __setitem__(self, key, val):
        if isinstance(key, str):
            key = self.__replaceColName(key)
        else:
            key += self._zOff
        dict.__setitem__(self, key, val)

    ##
    # Converts all of the keys that are strings to
    # integers, assuming that strings are xl colnames.
    # unnecessary function only for client side conviencence.
    def __ReplaceColName(self, k):
        place = 1
        index = 0
        for c in reversed(k):
            index += place * (int(c, 36) - 9)
            place *= 26
        index -= 1
        return index


###############################################################################
#
# Data Mapping - Maybe merge into mpCmd
#
###############################################################################


##
# Source of my woes
# returns $f(\vec{r})$
# unless f is a string then it returns f
# or if f is an int then it returns r[f]
# Need a better structure for handeling functions
def evaluate(f, r):
    if isinstance(f, str):
        return f
    elif isinstance(f, int):
        return r[f]  # if zeroOffset else r[f-1] // that should be done before
    else:  # this looks like horrible recursion
        return f[0](*[evaluate(a, r) for a in f[1]])


##
# performs [M].[F] = B
# where F_j(M_i) = B[i][j]
# default is by cols of F and rows of M
# M must be transposed beforehand for other method
def mapOperate(M, D):
    X = list(range(max(D.keys())))
    Y = list(range(len(M[0])))
    B = [[None for i in X] for j in Y]
    for i, r in zip(X, M):
        for k in zip(D):
            B[i][k] = evaluate(D[k], r)
    return B


###############################################################################
#
# User Accessable Functions
#
###############################################################################


def xlmp(cmd, mapByCol=True, fBook='', tBook='', fSheet=0, tSheet='xlmp',
         fCrng=(0, -1), fRrng=(0, -1), tr0=0, tc0=0):
    xlrw = ixl(readByRow=mapByCol)
    M = xlrw.guessRead(fBook, fSheet, fCrng, fRrng)
    M = mapOperate(M, cmd)
    xlrw.guessWrite(M, tBook, tSheet, tr0, tc0)


def genSubSheetsByIds(M, idLocs):
    idFunc = lambda x: [x[d] for d in idLocs]
    M.sort(key=idFunc)
    return [list(m) for k, m in groupby(M, idFunc)]


def xlSubMapping(subCmd, subSheets):
    pass