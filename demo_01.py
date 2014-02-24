from colMapper import *
import xlrd
import xlwt
fwb = xlrd.open_workbook('demo_1.xls')
fws = fwb.sheet_by_index(0)
twb = xlwt.Workbook()
tws = twb.add_sheet('Sheet 1')

demoCmd = MapCmd({
    "A": opList([mapIs, "Hello, World"]),
    "B": opList([mapSum, "A", "B", "C"]),
    "C": opList([mapAss, "C", "D", "E"]),
    "D": opList([mapProd, opList([mapSum, "B", "A"]), "A", "B"]),
    "E": "F"
})

interpColMap(demoCmd, fws, tws, 1, 0, 1)
twb.Save('out_demo_1.xls')