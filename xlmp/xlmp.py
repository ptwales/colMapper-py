#mapRow = False
#mapCol = True
byCol_not_byRow = True


### python-excel dependencies
##
# ONLY XLRD Dependency
# creates matrix of values of xlrd.Sheet object
def pullSheet(s, xrng=(0, -1), yrng=(0, -1)):
    assert not s.ragged_rows
    (crng, rrng) = (xrng, list(yrng)) if byCol_not_byRow else (yrng, list(xrng))
    if rrng[1] < 0:
        rrng[1] += s.nrows
    M = [s.row_values(r, *crng for r in range(*rrng)
    return (M if byCol_not_byRow else zip(*M))


##
# ONLY XLWT dependency
# writes maxtrix M to xlwt.Sheet object
# offset by r0 and c0
def writeSheet(s, M, p=(0, 0)):
    for i in p:
        assert i >= 0
    if not byCol_not_byRow:
        M = zip(*M)
        p = p[::-1]
    for x in range(len(M)):
        for y in range(len(M[x])):
            s.write(p[0] + x, p[1] + y, M[x][y])
### END python-excel dependencies


##
# Converts all of the keys that are strings to 
# integers, assuming that strings are xl colnames.
# unnecessary function only for client side conviencence.
def __ReplaceColNames(D):
    for k in list(D.keys()):
        if type(k) is str:
            place = 1
            index = 0
            for c in reversed(k):
                index += place * (int(c, 36) - 9)
                place *= 26  # Grossest part is that this executes after we are done
            index -= 1
        if k != index:
            D[index] = D.pop(k)

    

##
# M might be a smaller matrix than the range on the sheet
# it is to print to.  _EG_ if we writing to columns, 0, 1, and 3,
# then we M is will be of dim (:,3) and needs to become (:,4) with
# a blank column 2
def __insertNullCols(M, D):
    assert len(M) == len(D)
    B = [[None for y in range(max(D.keys()))] for x in range(len(M[0]))]
    for m, k in zip(zip(*M), D):
        B[k] = m
    return zip(*B)


##
# Source of my woes
# returns $f(\vec{r})$
# unless f is a string then it returns f
# or if f is an int then it returns r[f]
# Need a better structure for handeling functions
def __evalMapCmd(f, r):
    if type(f) is str:
        return f
    elif type(f) is int:
        return r[f] # if zeroOffset else r[f-1]
    else:  # this looks like horrible recursion
        return f[0](*[__evalMapCmd(a, r) for a in f[1]])


##
# performs [M].[F] = B
# where F_j(M_i) = B[i][j]
# default is by cols of F and rows of M
# M must be transposed beforehand for other method
def __mMap(M, F):
    B = __insertNullCols([__evalMapCmd(f, m) for f in F] for m in M], F)


##
# Maps from fSheet to tSheet according to Cmd
def xlmap(Cmd, fSheet, tSheet, mapX=byCol_not_byRow,
          frng=(0, -1), tp=(0, 0)):
    global byCol_not_byRow
    byCol_not_byRow = mapX
    if mapX:
        __ReplaceColNames(Cmd)
    writeSheet(tSheet, 
               __mMap(pullSheet(fSheet, frng), 
                      [Cmd[k] for k in Cmd]),
               tp)
    # M = pullSheet(fSheet, frng)
    # M = __mMap(M, [Cmd[k] for k in Cmd])
    # writeSheet(tSheet, M, tp)
