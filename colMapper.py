# python modules for spreadsheet manipulating not loaded
# CONTAINS PSEUDO CODE FOR IT
# => pseudo code is VBA

"""
Use Dictionary or Queue?

mapCmds: dict of lists of more lists or functions or parameters
first element in list is the function, rest are parameters (Polish/LISP)

dict key is column to map to
Total evaluation of a dict element contains the map to a column

for each element in dict (or queue
    for each row
        recurslively evaluate the list.
            first element is function,
            following are arguements
                if arguement is a list
                    evaluate that list
                else arguement is a column
                    get value of that column for current row.
        returned value becomes the value of the (current row, dict key column)
"""
        
class colMapCmd(object):

    def __init__(self):
        # declare dict
    
    def addMapCmd(self, IMapOper, destColumn):
    
    def stackMapCmd(self, IMapOper):
    

# by row
def colMap(mapCommands, fromSheet, toSheet, topRow=2, bottomRow=0, toTopRow=2):
    if bottomRow < 1:
        bottomRow = fromSheet.UsedRange.Rows.Count
    assert topRow < bottomRow and toTopRow > 0
    for row in range(topRow, bottomRow):
        for key in mapCommands.keys():
            toSheet.Cells(row, key).value = evaluateCommand(mapCommands[key], fromSheet)

def evaluateCommand(mapCmd, fromSheet, row):
    if len(mapCmd) = 1:
        return fromSheet.Cells(row, mapCmd[0]).Value
    else:
        args[]
            for i in range(1,len(mapCmd)):
                args[i] = evaluateCommand(mapCmd[i], fromSheet, row)
        return mapCmd[0](args)


