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


def evalOpList(opList, fromSheet, row):
    if type(opList) is str:
        return opList
    elif type(opList) is int:
        return fromSheet.cell(row, opList).value
    else:
        itOpList = iter(opList)
        next(itOpList)
        return opList[0](*[evalOpList(it, fromSheet, row) for it in itOpList])


def interpColMap(mapCmd, fmSheet, toSheet, rTop=0, rBot=0, rToTop=0):
    if rBot < 0:
        rBot = fmSheet.nrows
    __MapCmdConvert(mapCmd)
    assert rTop < rBot and rToTop >= 0
    for row in range(rTop, rBot):
        for key in list(mapCmd.keys()):
            toSheet.write(row, key, evalOpList(mapCmd[key], fmSheet, row))
