#mapRow = False
#mapCol = True
map_byCol_not_byRow = True


### python-excel dependencies
### mapCol and mapRow should only be in these two
### So it should read it transposed and write it transposed
##
# ONLY XLRD Dependency
# creates matrix of values of xlrd.Sheet object
def pullSheet(s, xr=(0,-1), yr=(0,-1)):
    assert not s.ragged_rows
    rect = (xr, yr)
    if not map_byCol_not_byRow:
        rect = reversed(rect)
    if rect[0][1] < 0:
        rect[0][1] += s.nrows
    M = []
    for r in range(*rect[0]):
        M.append(s.row_values(r, *rect[1]))
    if map_byCol_not_byRow:
        return M
    else:
        return zip(*M)


##
# ONLY XLWT dependency
# writes maxtrix M to xlwt.Sheet object
# offset by r0 and c0
def writeSheet(s, M, p=(0,0)):
    for i in p:
        assert i >= 0
    if not map_byCol_not_byRow:
        M = zip(*M)
        p = reversed(p)
    for x in range(len(M)):
        for y in range(len(M[x])):
            s.write(p[0] + x, p[1] + y, M[x][y])
### END python-excel dependencies


##
# M might be a smaller matrix than the range on the sheet
# it is to print to.  _EG_ if we writing to columns, 0, 1, and 3,
# then we M is will be of dim (:,3) and needs to become (:,4) with
# a blank row in column 2
def __insertNullRows(M, D):
    assert len(M) == len(D)
    B = [[None for y in range(max(D.keys()))] for x in range(len(M[0]))]
    for r, k in zip(range(len(M)), D):
        B[k] = M[r]
    return B


def __evalMapCmd(f, r):
    if type(f) is str:
        return f
    elif type(f) is int:
        return r[f]
    else:
        # this looks like horrible recursion
        return f[0](*[__evalMapCmd(a, r) for a in f[1]])


##
# performs [M].[F] = B
# where F_j(M_i) = B[i][j]
# default is by cols of F and rows of M
# M must be transposed beforehand for other method
def __mMap(M, F):
    B = []
    for r in M:
        B.append([])
        for f in F:
            B[-1].append( __evalMapCmd(f, r))
    return B


def xlmap(Cmd, fSheet, tSheet, mapX=map_byCol_not_byRow,
          frng=(0,-1), tp=(0,0)):
    global map_byCol_not_byRow
    map_byCol_not_byRow = mapX
    M = pullSheet(fSheet, frng)
    if mapX:
        #__ReplaceColNames(Cmd)
        pass
    M = __mMap(M, [Cmd[k] for k in Cmd])
    M = __insertNullRows(M, Cmd)
    writeSheet(tSheet, M, tp)
