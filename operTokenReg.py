tokenReg = {
        '=': mapIs,
        '==': mapAss,
        '+': mapSum,
        '*': mapProd
        }
def getTokenReg(token):
    return tokenReg.get(token)
