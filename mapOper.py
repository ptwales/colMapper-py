__FAIL__ = "~FAIL!"
"""
def __stripNULL__(valueList):
    for value in valueList:
        if value = NULL:
            valueList.remove(value)
""" 
tokenReg = {
    '=': mapIs,
    '==': mapAss,
    '+': mapSum,
    '*': mapProd
    }
        
def getTokenReg(token):
    return tokenReg.get(token)

def mapIs(valueList):
    return values[0]

def mapSum(valueList):
    return sum(valueList)
    
def mapAss(valueList):
    #__stripNull__(valueList)
    if len(valueList) == 1:
            return valueList[0]
        for i in range(0, len(valueList) - 1):
            if valueList[i] != valueList[i+1]:
                return __FAIL__
        return valueList[0]

def mapCond(valueList):
    # valueList = (BOOL, TRUE_VALUE, FALSE_VALUE)
    assert len(valueList) = 3
    return valueList[1] if valueList[0] else valueList[2]

def mapProd(valueList):
    #__stripNull__(valueList)
    return reduce(operator.mul, valueList, 1)

