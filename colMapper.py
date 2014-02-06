from mmap import mmap
from xlrd import open_workbook

"""
mapCmds: dict of lists of more lists or functions or parameters
first element in list is the function, rest are parameters (Polish/LISP)

dict key is column to map to
Total evaluation of a dict element contains the map to a column

"""
    
class colMapCmd(object):

    def __init__(self, parent, *cols):
        self.parent = parent
        self.cols = cols                # can be another colMapCmd

    def evaluate(self, fromSheet, row):
        if len(self.cols) = 1:
            return fromSheet.cell(row, self.cols[0]).value
        else:
            args[]
                for i in range(1,len(self.cols)):
                    args[i] = self.cols[i].evaluate(fromSheet, row)
            return self.cols[0](args)


class sheetMapCmd(dict):

    def __init__(self, parent, *colMap):
        self.parent = parent
        self._ = *colMap

    def validate(self):
        """
        assert that every key be a key to a sheet column
        therefore it must be either
          a string of no more than 3 letters 
          or less than (max number of excel rows)

        will fail if keys are not colMapCmd
        """


# by row
def interpColMap(mapCommands, fromSheet, toSheet, topRow=2, bottomRow=0, toTopRow=2):
    if bottomRow < 1:
        bottomRow = fromSheet.nrows
    assert topRow < bottomRow and toTopRow > 0
    for row in range(topRow, bottomRow):
        for key in mapCommands.keys():
            toSheet.cell(row, key).value = evaluateCommand(mapCommands[key], fromSheet)

