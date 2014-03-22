#import xlutils

mapCols = True
mapRows = False
EXCEL2010_MAXCOL = 16372
EXCEL2008_MAXCOL = 256
MAX_COL = EXCEL2008_MAXCOL


# this might make things so much easier or not
def __pullSheet(s, r0=0, c0=0, rf=-1, cf=-1):
    if rf == -1:
        rf = s.nrows
    if cf == -1:
        cf = s.ncols
    assert r0 < rf and c0 < cf
    x0 = 0
    y0 = 0
    xf = rf - r0
    yf = cf - c0
    m = [[0 for y in range(y0, yf)] for x in range(x0, xf)]
    for x, r in list(range(x0, xf)), list(range(r0, rf)):
        for y, c in list(range(y0, yf)), list(range(c0, cf)):
            m[x][y] = s.cells(r, c).value
    return m

'''
colAtoInt
accepts a column index string in typical excel format
converts it to the appropriate integer.

xlrd supplies colIndex to colname but not colname to colIndex
Original Author
https://stackoverflow.com/questions/763691/
programming-riddle-how-might-you-translate-an-excel-column-name-to-a-number
'''


#
def __colName_to_colIndex(colName):
    index = 0
    if type(colName) is str:
        assert len(colName) <= 3
        place = 1
        for c in reversed(colName):
            index += place * (int(c, 36) - 9)
            place *= 26  # Grossest part is that this executes after we are done
        index -= 1
    assert index <= MAX_COL  # check if larger than column count
    return index


#
def __MapCmdConvert(mapCmd):
    for k in list(mapCmd.keys()):
        colIdx = __colName_to_colIndex(k)  # convert to integer
        if k != colIdx:
            mapCmd[colIdx] = mapCmd.pop(k)


def __accessor(mapBy):
    if mapBy:
        return (lambda r, k: (r, k))
    else:
        return (lambda c, k: (k, c))


#Playing with byRow
def __evalOps(Ops, fws, i, a):
    if type(Ops) is str:
        return Ops
    elif type(Ops) is int:
        return fws.cell(*a(i, Ops)).value
    else:
        return Ops[0](*[__evalOps(o, fws, i, a) for o in Ops[1]])


def interpMap(mapCmd, fmSheet, toSheet, mapBy=mapCols,
              from_start=0, from_end=-1, to_start=0):
    if from_end == -1:
        from_end = fmSheet.nrows if mapBy else fmSheet.ncols
    assert from_start < from_end and to_start >= 0
    if mapBy:
        __MapCmdConvert(mapCmd)
    a = __accessor(mapBy)
    for i in range(from_start, from_end):
        for key in list(mapCmd.keys()):
            toSheet.write(*a(i, key + [__evalOps(mapCmd[key], fmSheet, i, a)]))


'''
def genSubSheetsByIds(fmSheet, idLocs, mapBy=mapCols,
                      from_start=0, from_end=-1):
    if from_end == -1:
        from_end = fmSheet.nrows if mapBy else fmSheet.ncols
    a = __accessor(mapBy)
    sub_start = from_start

    def getIDs(i):
        return [fmSheet.cells(*a(sub_start, loc)).value for loc in idLocs]

    sub_sheets = []

    def appendSheet(start, end):
        #PSEUDO I
        sub_sheets.append(xlutils.copy(fmSheet, sub_start, i - 1))

    crntIDs = getIDs(sub_start)
    for i in range(from_start, from_end):
        nowIDs = getIDs(i)
        if nowIDs != crntIDs:
            appendSheet(sub_start, i - 1)
            crntIDs = nowIDs
            sub_start = i
    appendSheet(sub_start, from_end)
    return sub_sheets
'''


#
def interpSubMapping(mapCmd, subSheets, toSheet, idLocs, mapBy=mapCols,
                     toStart=0):
    to_end = toStart
    for ss in subSheets:
        interpMap(mapCmd, ss, toSheet, (not mapBy), to_start=to_end)
        to_end += max(mapCmd.keys()) + 1



