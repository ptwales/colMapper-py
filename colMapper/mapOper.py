from operator import mul

__FAIL__ = "~FAIL!"

"""
def __stripNULL__(valueList):
    for value in valueList:
        if value = NULL:
            valueList.remove(value)
"""


def mapVal(valueList):
    return valueList[0]

def mapSum(valueList):
    return sum(valueList)


def mapAss(valueList):
    #__stripNull__(valueList)
    if len(valueList) == 1:
        return valueList[0]
    else:
        for i in range(0, len(valueList) - 1):
            if valueList[i] != valueList[i + 1]:
                return __FAIL__
        return valueList[0]


def mapCond(valueList):
    # valueList = (X == Y, TRUE_VALUE, FALSE_VALUE)
    assert len(valueList) == 4
    return valueList[2] if valueList[0] == valueList[1] else valueList[3]


def mapProd(valueList):
    #__stripNull__(valueList)
    return reduce(mul, valueList, 1)


tokenReg = {
    '=': mapVal,
    '$': mapVal,
    '==': mapAss,
    '+': mapSum,
    '*': mapProd,
    '?': mapCond,
    }


def getTokenReg(token):
    return tokenReg.get(token)
