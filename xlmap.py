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
def __mMap(M, F): # [M].[F] = [B]
    B = __prealloc(len(M), len(F))
    for M_r, B_r in zip(M, B):
        for B_el, f in zip(B_r, F):
            B_el = __evalMapCmd(f, M_r)
    return B
