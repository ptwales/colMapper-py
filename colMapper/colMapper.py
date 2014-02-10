from mmap import mmap
from xlrd import open_workbook

"""
mapCmds: dict of lists of more lists or functions or parameters
first element in list is the function, rest are parameters (Polish/LISP)

dict key is column to map to
Total evaluation of a dict element contains the map to a column

"""
    
class colMapCmd( object ):

    def __init__( self, *cols ):
        self._ = cols                # can be another colMapCmd

    def evaluate( self, fromSheet, row ):
        if len( self.cols ) = 1:
            return fromSheet.cell( row, self._.get(0) ).value
        else:
            args[]
                for i in range( 1, len( self._ ) ):
                    args[i] = self._.get( i ).evaluate( fromSheet, row )
            return self._.get(0)( args )


class sheetMapCmd( dict ):

    def __init__( self, dictColMapCmd ):
        self._ = { cmc.key : cmc.col for cmc in dicColMapCmd }
        if not self.__validate__:
            print( "dictColMapCmd is invalid" )
            del self._{:}
            
    def __validate__( self ):
        """
        assert that every key be a key to a sheet column
        therefore it must be either
          a string of no more than 3 letters 
          or less than (max number of excel rows)

        will fail if keys are not colMapCmd
        """
        return False


def interpColMap( mapCommands, fromSheet, toSheet, topRow=2, bottomRow=0, toTopRow=2 ):
    if bottomRow < 1:
        bottomRow = fromSheet.nrows
    assert topRow < bottomRow and toTopRow > 0
    for row in range( topRow, bottomRow ):
        for key in mapCommands.keys():
            toSheet.cell( row, key ).value = evaluateCommand( mapCommands.get( key ), fromSheet )

