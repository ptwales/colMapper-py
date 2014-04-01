from . import xlmp
from itertools import groupby

def genSubSheetsByIds(fmSheet, idLocs, mapX=xlmp.byCol_not_byRow,
                      frng=(0, -1)):
    global xlmp.byCol_not_byRow
    xlmp.byCol_not_byRow = mapX
    M = xlmp.pullSheet(fmSheet, xr=frng)
    idFunc = lambda x: [x[d] for d in idLocs]
    M.sort(key=idFunc)
    return [list(m) for k, m in groupby(M, idFunc)]
    


# define functions to create subSheets, which will be matricies.
def xlSubMapping(subCmd, subSheets, fSheet, tSheet, mapX=not xlmp.byCol_not_byRow,
                 tp=(0,0)):
    tx, ty = tp  # tuples are immuteable
    for m in subSheets:
        b = xlmp.__mMap(zip(*m), subCmd)
        xlmp.writeSheet(tSheet, (tx, ty))
        tx += len(b)  # and we are modifying tx or tp[0]
