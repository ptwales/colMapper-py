import colMapper

fwp = open_workbook( 'fromSheet.xls' )
twp = open_workbook( 'toSheet.xls' )

cmCmd = sheetMapCmd ( { "A": colMapCmd( [mapIs, "Hello, World"] ) # map is accepts only 1 arg
                      , "B": colMapCmd( [mapSum, "A", "B", "C"] )
                      , "C": colMapCmd( [mapAss, "A", "C", "D"] )
                      , "D": colMapCmd( [mapProd, colMapCmd( [mapSum, "A", "B"] ), "C", "D"] ) # stacking commands
                      } )

interpColMap( cmCmd, fwp, twp, 1, 0, 1 )
