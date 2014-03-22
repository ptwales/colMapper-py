def __prealloc(X, Y):
    return [[0 for y in range(Y)] for x in range(X)]


def pullSheet(s, r0=0, c0=0, rf=-1, cf=-1):
    if rf == -1:
        rf = s.nrows
    if cf == -1:
        cf = s.ncols
    assert r0 < rf and c0 < cf
    xf = rf - r0
    yf = cf - c0
    M = __prealloc(xf, yf)
    for x, r in zip(range(xf), range(r0, rf)):
        for y, c in zip(range(yf), range(c0, cf)):
            M[x][y] = s.cell(r, c).value
    return M


def writeSheet(s, B, r0=0, c0=0):
    for x in range(len(B)):
        for y in range(len(B[x])):  # assert len(B[r]) == len(D)
            s.write(r0 + x, c0 + y, B[x][y])


def __insertNullRows(M, D):
    B = __prealloc(max(D.keys()), len(M[0]))
    for r, k in zip(range(len(M)), D):
        B[k] = M[r]
    return B


mapRow = False
mapCol = True


def __evalMapCmd(f, r):
    if type(f) is str:
        return f
    elif type(f) is int:
        return r[f]
    else:
        # this looks like horrible recursion
        return f[0](*[__evalMapCmd(a, r) for a in f[1]])


# default is by cols of F and rows of M
# M must be transposed beforehand for other method
def __mMap(M, F):  # [M].[F] = [B]
    X = len(M)
    Y = len(F)
    B = __prealloc(X, Y)
    for r, x in zip(M, range(X)):
        for f, y in zip(F, range(Y)):
            B[x][y] = __evalMapCmd(f, r)
    return B


def xlmap(Cmd, fSheet, tSheet, mapX=mapCol,
          fStart=0, fStop=-1, tStart=0):
    if fStop == -1:
        fStop = fSheet.nrows if mapX else fSheet.ncols
    assert fStart < fStop
    if mapX:
        M = pullSheet(fSheet, r0=fStart, rf=fStop)
        #__ReplaceColNames(Cmd)
    else:
        M = pullSheet(fSheet, c0=fStart, cf=fStop)
        # transpose M
    M = __mMap(M, [Cmd[k] for k in Cmd])
    M = __insertNullRows(M, Cmd)
    if mapX:
        writeSheet(fSheet, M, r0=tStart)
    else:
        # transpose M
        writeSheet(fSheet, M, c0=tStart)