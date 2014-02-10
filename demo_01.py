import colMapper

fwb = open_workbook( 'fromSheet.xls' )
fws = fwb.Sheets( 1 )
twb = open_workbook( 'toSheet.xls' )
twb = twb.Sheets( 1 )

demoCmd = MapCmd ( { "A": opList( (mapIs, "Hello, World") ) # map is accepts only 1 arg
                   , "B": opList( (mapSum, "A", "B", "C") )
                   , "C": opList( (mapAss, "A", "C", "D") )
                   , "D": opList( (mapProd, opList( (mapSum, "A", "B") ), "C", "D") ) # stacking commands
                 } )

interpColMap( demoCmd, fws, tws, 1, 0, 1 )

# Pseudo Code
fwb.close
twb.save
twb.close
