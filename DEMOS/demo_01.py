# import the files

fwp = open_workbook( 'fromSheet.xls' )
twp = open_workbook( 'toSheet.xls' )

cmCmd = sheetMapCmd ( { "A": [mapIs, "Hello, World"]
                      , "B": [mapSum, "A", "B", "C"]
                      , "C": [mapAss, "C", "D"]
                      , "D": [mapProd, "B", "E"]
                      } )
