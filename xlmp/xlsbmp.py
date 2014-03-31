from . import xlmp


def genSubSheetsByIds(fmSheet, idLocs, mapX=xlmp.byCol_not_byRow,
                      frng=(0, -1)):
    global xlmp.byCol_not_byRow
    xlmp.byCol_not_byRow = mapX
    M = xlmp.pullSheet(fmSheet, xr=frng)
    for d in idLocs:
        M.sort(key=lambda el: el[d])
        
    def __matchRowByIDs(ra, rb, idLocs):
        return all([ra[d] == rb[d] for d in idLocs])
        
    m = [[M[0]]]
    for r in M[1:]:
        if __matchRowByIDs([m[-1][-1], r[d], idLocs):
            m[-1].append(r)
        else:
            m.append([r])
    
    return m


# define functions to create subSheets, which will be matricies.
def xlSubMapping(subCmd, subSheets, fSheet, tSheet, mapX=not xlmp.byCol_not_byRow,
                 tp=(0,0)):
    for m in subSheets:
        b = xlmp.__mMap(zip(*m), subCmd)
        xlmp.writeSheet(tSheet, tp)
        tp[0] += len(b)
