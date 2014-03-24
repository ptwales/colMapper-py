from . import xlmp


def genSubSheetsByIds(fmSheet, idLocs, mapX=mapCols,
                      fStart=0, fStop=-1):
    if mapX:
        M = xlmp.pullSheet(fmSheet, r0=fStart, rf=fStop)
    else:
        M = xlmp.__transpose(xlmp.pullSheet(fmSheet, c0=fStart, cf=fStop))
    # sort M by idLocs
    m = [[[None]]]
    m_i = 0
    for i in range(0, len(M)):
        pass
    return m


# define functions to create subSheets, which will be matricies.
def xlSubMapping(subCmd, subSheets, fSheet, tSheet, mapX=mapRow,
                 tStart=0):
    if mapX:
        pass
    else:
        pass
