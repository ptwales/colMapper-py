import xlrd
import xlwt

byCol_not_byRow = True

###############################################################################
#
# XLRD XLWT Interfaces
#
###############################################################################


##
# creates matrix of values of xlrd.Sheet object
def __pullSheet(s, xrng=(0, -1), yrng=(0, -1)):
    assert not s.ragged_rows
    (crng, rrng) = (xrng, list(yrng)) if byCol_not_byRow else (yrng, list(xrng))
    if rrng[1] < 0:
        rrng[1] += s.nrows
    M = [s.row_values(r, *crng) for r in range(*rrng)]
    return (M if byCol_not_byRow else zip(*M))


##
# writes maxtrix M to xlwt.Sheet object
# offset by r0 and c0
def __writeSheet(M, s, p=(0, 0)):
    for i in p:
        assert i >= 0
    if not byCol_not_byRow:
        M = zip(*M)
        p = p[::-1]
    for x in range(len(M)):
        for y in range(len(M[x])):
            s.write(p[0] + x, p[1] + y, M[x][y])


def __readBook(book_path, sheet_index=0, xrng=(0, -1), yrng=(0, -1)):
    book = xlrd.open_workbook(book_path)
    try:
        sheet = book.sheet_by_index(sheet_index)
    except TypeError:
        sheet = book.sheet_by_name(sheet_index)
    return __pullSheet(sheet, xrng, yrng)


def __writeBook(M, book_path, sheet_name='xlmp', p=(0, 0)):
    book = xlwt.Workbook()
    sheet = book.add_sheet(sheet_name)
    __writeSheet(M, sheet, p)
    book.save(book_path)


###############################################################################
#
# Optional/User-Conveinece Interperters
#  Removes zero offsets allow Excel Column names
#
###############################################################################


##
# Converts all of the keys that are strings to
# integers, assuming that strings are xl colnames.
# unnecessary function only for client side conviencence.
def __ReplaceColNames(D):
    for k in list(D.keys()):
        if isinstance(k, str):
            place = 1
            index = 0
            for c in reversed(k):
                index += place * (int(c, 36) - 9)
                place *= 26
            index -= 1
        if k != index:
            D[index] = D.pop(k)

'''
def __ReplaceZeroOffset(D):
'''


###############################################################################
#
# Data Mapping
#
###############################################################################


##
# M might be a smaller matrix than the range on the sheet
# it is to print to.  _EG_ if we writing to columns, 0, 1, and 3,
# then we M is will be of dim (:,3) and needs to become (:,4) with
# a blank column 2
def __insertNullCols(M, D):
    assert len(M) == len(D)
    B = [[None for y in range(max(D.keys()))] for x in range(len(M[0]))]
    for m, k in zip(zip(*M), D):
        B[k] = m
    return zip(*B)


##
# Source of my woes
# returns $f(\vec{r})$
# unless f is a string then it returns f
# or if f is an int then it returns r[f]
# Need a better structure for handeling functions
def __evalMapCmd(f, r):
    if isinstance(f, str):
        return f
    elif isinstance(f, int):
        return r[f]  # if zeroOffset else r[f-1] // that should be done before
    else:  # this looks like horrible recursion
        return f[0](*[__evalMapCmd(a, r) for a in f[1]])


##
# performs [M].[F] = B
# where F_j(M_i) = B[i][j]
# default is by cols of F and rows of M
# M must be transposed beforehand for other method
def __mMap(M, F):
    return __insertNullCols([[__evalMapCmd(f, m) for f in F] for m in M], F)


###############################################################################
#
# User Accessable Functions
#
###############################################################################


##
# Maps from fBook to tBook according to Cmd
def xlmap(Cmd, fBook='', tBook='', fSheet=0, tSheet='xlmp',
          mapX=byCol_not_byRow, frng=(0, -1), tp=(0, 0)):
    global byCol_not_byRow
    byCol_not_byRow = mapX
    if mapX:
        __ReplaceColNames(Cmd)
    try:
        M = __pullSheet(fSheet, frng)
    except AttributeError:
        M = __readBook(fBook, fSheet, frng)
    M = __mMap(M, [Cmd[k] for k in Cmd])
    try:
        __writeSheet(M, tSheet, tp)
    except AttributeError:
        __writeBook(M, tBook, tSheet, tp)
