from . import xlmp


def genSubSheetsByIds(fmSheet, idLocs, mapX=xlmp.mapCols,
                      fStart=0, fStop=-1):
    if mapX:
        M = xlmp.pullSheet(fmSheet, r0=fStart, rf=fStop)
    else:
        M = xlmp.__transpose(xlmp.pullSheet(fmSheet, c0=fStart, cf=fStop))
    for d in idLocs:
        M.sort(key=lambda el: el[d])
    m = [[M[0]]]
    for r in M[1:]:
        if all([m[-1][-1][d] == r[d] for d in idLocs]):
            m[-1].append(r)
        else:
            m.append([r])
    return m


# define functions to create subSheets, which will be matricies.
def xlSubMapping(subCmd, subSheets, fSheet, tSheet, mapX=xlmp.mapRow,
                 tStart=0):
    for m in subSheets:
        b = xlmp.__mMap(zip(*m), subCmd)
        if mapX:
            xlmp.writeSheet(tSheet, r0=tStart)
        else:
            xlmp.writeSheet(tSheet, c0=tStart)
        tStart += len(b)