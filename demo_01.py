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
    'B': [mapOper.mapSum, 1, 2, 3],
    'C': [mapOper.mapAss, 4, 5, 6],
    'D': [mapOper.mapProd, [mapOper.mapSum, 2, 1], 1, 2],
    'E': 6
    }

colMapper.interpColMap(demoCmd, fws, tws, 1, 0, 1)
twb.save('out_demo_1.xls')
