mapRow = False
mapCol = True
#map_byCol_not_byRow = True


##
# Preallocation of Matrix M, assures that it is of the correct dimensions
# probably poor form
def __prealloc(X, Y):
    return [[None for y in range(Y)] for x in range(X)]


### python-excel dependencies
### mapCol and mapRow should only be in these two
### So it should read it transposed and write it transposed
##
# ONLY XLRD Dependency
# creates matrix of values of xlrd.Sheet object
def pullSheet(s, r0=0, c0=0, rf=-1, cf=-1):
    assert not s.ragged_rows
    if rf < 0:
        rf += s.nrows
    if cf < 0:
        cf += s.ncols
    M = [[]]
    for r in range(r0, rf):
        M.append(s.row_values(r, c0, cf))
    return M


##
# ONLY XLWT dependency
# writes maxtrix M to xlwt.Sheet object
# offset by r0 and c0
def writeSheet(s, M, r0=0, c0=0):
    assert r0 >= 0 and c0 >= 0
    for x in range(len(M)):
        for y in range(len(M[x])):
            s.write(r0 + x, c0 + y, M[x][y])


### END python-excel dependencies

##
# M might be a smaller matrix than the range on the sheet
# it is to print to.  _EG_ if we writing to columns, 0, 1, and 3,
# then we M is will be of dim (:,3) and needs to become (:,4) with
# a blank row in column 2
def __insertNullRows(M, D):
    assert len(M) == len(D)
    B = __prealloc(max(D.keys()), len(M[0]))  # only place I should need __prealloc
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
    X = len(M)
    Y = len(F)
    B = __prealloc(X, Y)
    for r, i in zip(M, range(X)):  # should not need __prealloc here
        for f, j in zip(F, range(Y)):
            B[i][j] = __evalMapCmd(f, r)
    return B


##
# transpose matrix M
def __transpose(M):
    return zip(*M)


def xlmap(Cmd, fSheet, tSheet, mapX=mapCol,
          fStart=0, fStop=-1, tStart=0):
    if mapX:  # move this to pullSheet and writeSheet
        M = pullSheet(fSheet, r0=fStart, rf=fStop)
        #__ReplaceColNames(Cmd)
    else:
        M = __transpose(pullSheet(fSheet, c0=fStart, cf=fStop))
    M = __mMap(M, [Cmd[k] for k in Cmd])
    M = __insertNullRows(M, Cmd)
    if mapX:
        writeSheet(tSheet, M, r0=tStart)
    else:
        M = __transpose(M)
        writeSheet(tSheet, M, c0=tStart)
