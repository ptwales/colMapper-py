class colMapper(object):
    
    def __init__(self):
        # declare queue 

    # Set map commands
    def setMapCmdWithString(self, stringCmd):

    def queueMapCmd(self, IMapOper):
    
    def doMapCmd(self, IMapOper):

    # Engine
    def mapTo(self, fSheet, tSheet, topRow=2, bottomRow=0, toTopRow=2):
        if bottomRow < 1:
            bottomRow = fSheet.UsedRange.Rows.Count
        assert topRow < bottomRow and toTopRow > 0
        for row in range(topRow, bottomRow):
            for command in cmdQueue:
                self.doMapCmd(command)
