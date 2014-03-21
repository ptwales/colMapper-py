EXCEL2010_MAXCOL = 16372  # is this in xlrd/xlwt?
EXCEL2008_MAXCOL = 256
MAX_COL = EXCEL2008_MAXCOL

"""
colAtoInt
accepts a column index string in typical excel format
converts it to the appropriate integer.

xlrd supplies colIndex to colname but not colname to colIndex
Original Author
https://stackoverflow.com/questions/763691/
programming-riddle-how-might-you-translate-an-excel-column-name-to-a-number
"""
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


def __MapCmdConvert(mapCmd):
    for k in list(mapCmd.keys()):
        colIdx = __colName_to_colIndex(k)  # convert to integer
        if k != colIdx:
            mapCmd[colIdx] = mapCmd.pop(k)


'''
@author Philip Wales

@param Ops 
@param fromSheet
@param r
'''
def __evalOps(Ops, fws, r):
    if type(Ops) is str:
        return Ops
    elif type(Ops) is int:
        return fws.cell(r, Ops).value
    else:
        return Ops[0](*[__evalOps(o, fws, r) for o in Ops[1]])

'''
@author Philip Wales

@param
'''
def interpColMap(mapCmd, fmSheet, toSheet, rTop=0, rBot=0, rToTop=0):
    if rBot < 0:
        rBot = fmSheet.nrows  # somehow make this default arg
    __MapCmdConvert(mapCmd)   # convert all colNames to colIndex
    assert rTop < rBot and rToTop >= 0
    for row in range(rTop, rBot):
        for key in list(mapCmd.keys()):
            toSheet.write(row, key, __evalOps(mapCmd[key], fmSheet, row))

'''Playing with byRow
mapCols=True
mapRows=False


def __evalOps(Ops, fws, i, a):
    if type(Ops) is str:
        return Ops
    elif type(Ops) is int:
        return fws.cell(*a(i, Ops)).value
    else:
        return Ops[0](*[__evalOps(o, fws, i, a) for o in Ops[1]])


def interpMap(mapCmd, fmSheet, toSheet, mapBy=mapCols,
              from_start=0, from_end=-1, to_start=0):
    if mapBy:  # mapBy == mapCols
        if from_end < 0: from_end = fmSheet.nrows
        a = (lambda r, k: r, k)
        __MapCmdConvert(mapCmd)
    else: # mapBy == mapRows 
        if from_end < 0: from_end = fmSheet.ncols
        a = (lambda c, k: k, c)
        # no such thing as rowNames
    assert from_start < from_end and to_start >= 0
    for i in range(from_start, from_end):
        for key in list(mapCmd.keys()):
            toSheet.write(*a(i, key), __evalOps(mapCmd[key], fmSheet, i, a))
'''

def interpSubMapping(mapCmd, fmSheet, toSheet, idLocs, mapBy=mapCols,
                     from_start=0, from_end=-1, to_start=0):
    pass
'''
    def genSubSheets():
        yield ss
        
    sub_sheets = genSubSheets()
    to_end = to_start
    for ss in sub_sheets:
        interpMap(mapCmd, ss, toSheet, !mapBy, to_start:=to_end)
        to_end = toSheet.
'''          

# this might make things so much easier   
def pullSheet(s, r0=0, c0=0, rf=-1, cf=-1):
    if rf == -1: rf = s.nrows
    if cf == -1: cf = s.ncols
    assert r0 < rf and c0 < cf
    x0 = 0: y0 = 0 
    xf = rf - r0
    yf = cf - c0
    m = [[0 for y in range(y0, yf)] for x in range(x0, xf]
    for x, r in range(x0, xf), range(r0, rf):
        for y, c in range(y0, yf), range(c0, cf):
            m[x][y] = s.cells(r,c).value
    return m
    
