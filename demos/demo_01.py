from colMapper import *
from collections import defaultdict
import xlrd
import xlwt

fwb = xlrd.open_workbook('demo_1.xls')
fws = fwb.sheet_by_index(0)
twb = xlwt.Workbook()
tws = twb.add_sheet('Sheet 1')


def myFunc(x, y):
    return x**y
    
    
demoCmd = defautldict(tuple)   
demoCmd = {
# colName/colIndex : functions,
'A': 'I love Spam with eggs!',  # Strings will map that value
'B': 2,                         # Ints are zero-offset column indexes from fws
 2 : 3,                         # Column indexes are also zero offset
'D': (Sum, [0, 1, 2]),          # Column D is sum of columns A, B, C
'E': (Sum, [0, 1, (int, '99')]) # Escape Strings to Integers like this.
'F': (Sum, [0, 1, (mapOper.mapProd, 2, 3)]),  # Nesting is allowed
'G': (myFunc, [2,3]),           # Use your own functions
'H': ((lambda x, y: x**y), [2, 3]),  # Lambdas are ok too.
'I': (mapOper.mapAssert, [4, 5, 6])  # Some neat functions are provided.
}



colMapper.interpColMap(demoCmd, fws, tws)
twb.save('out_demo_1.xls')
