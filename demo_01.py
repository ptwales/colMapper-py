from colMapper import *
from collections import defaultdict
import xlrd
import xlwt

fwb = xlrd.open_workbook('demo_1.xls')
fws = fwb.sheet_by_index(0)
twb = xlwt.Workbook()
tws = twb.add_sheet('Sheet 1')

demoCmd = defaultdict(
    "A": [mapIs, "Hello, World"],
    "B": [mapSum, 1, 2, 3],
    "C": [mapAss, 3, 4, 5],
    "D": [mapProd, [mapSum, 2, 1], 1, 2],
    "E": 6
    )

interpColMap(demoCmd, fws, tws, 1, 0, 1)
twb.Save('out_demo_1.xls')
