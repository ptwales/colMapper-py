from colMapper import *
from collections import defaultdict
import xlrd
import xlwt

fwb = xlrd.open_workbook('demo_1.xls')
fws = fwb.sheet_by_index(0)
twb = xlwt.Workbook()
tws = twb.add_sheet('Sheet 1')

demoCmd = defaultdict(list)
demoCmd = {
    'A': [mapOper.mapVal, 'Hello, World'],
    'B': [mapOper.mapSum, 0, 1, 2],
    'C': [mapOper.mapAssert, 3, 4, 5],
    'D': [mapOper.mapProd, [mapOper.mapSum, 1, 0], 0, 1],
    'E': 5
    }

colMapper.interpColMap(demoCmd, fws, tws)
twb.save('out_demo_1.xls')
