#from mmap import mmap
#from xlrd import open_workbook
#from xlwt import Workbook

"""
mapCmds: dict of lists of more lists or functions or parameters
first element in list is the function, rest are parameters (Polish/LISP)

dict key is column to map to
Total evaluation of a dict element contains the map to a column

"""
EXCEL2010_MAXCOL = 16372 #is this in xlrd/xlwt?
EXCEL2008_MAXCOL = 256
MAX_COL = EXCEL2008_MAXCOL

"""
colAtoInt
accepts a column index string in typical excel format 
converts it to the appropriate integer.

xlrd supplies colIndex to colname but not colname to colIndex
Original Author
https://stackoverflow.com/questions/763691/programming-riddle-how-might-you-translate-an-excel-column-name-to-a-number
"""
def colName_to_colIndex(colName):
    index = 0
    if type(colName) is str:
        assert len(colName) <= 3
        place = 1
        for c in reversed(colName): #Original used [::-1]. I heard this was faster. Should test it
            index += place*(int(c, 36) - 9)
            place *= 26 # Grossest part is that this executes after we are done
    elif type(colName) is int:
        # nothing
    else:
        #raise error
    assert index <= MAX_COL # check if larger than column count
    return index
    
    
def MapCmdConvert(mapCmd):
    for k in mapCmd.keys():
        colIdx = colName_to_colIndex(k): # convert to integer
        if k != colIdx:
            mapCmd[colIdx] = mapCmd.pop(k)
            
            
def evalOpList(opList, fromSheet, row):
    if opList.count() == 1:
        return fromSheet.cell(row, self._.get(0)).value
    else:
        args = list()
        itOpList = iter(opList)
        next(itOpList)
        args = [evalOpList(it, fromSheet, row) for it in itOpList]
        return opList.get(0)(args)
        
        
def interpColMap(mapCmd, fmSheet, toSheet,
                 topRow=2, bottomRow=0, toTopRow=2):
    if bottomRow < 1:
        bottomRow = fmSheet.nrows
    MapCmdConvert(mapCmd)
    assert topRow < bottomRow and toTopRow > 0
    for row in range(topRow, bottomRow):
        for key in list(mapCmd.keys()):
            toSheet.write(row, key, evalOpList(mpCmd[key], fmSheet, row))
