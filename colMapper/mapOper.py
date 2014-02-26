from operator import mul

__FAIL__ = "~FAIL!"


def stripNull(valueList):
    valueList = [notNull for notNull in valueList if notNull]


def mapVal(valueList):
    assert len(valueList) == 1
    return valueList[0]


def mapSum(valueList):
    return sum(valueList)


def mapAss(valueList):
    stripNull(valueList)
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
    stripNull(valueList)
    return reduce(mul, valueList, 1)


tokenDict = {
    '$': mapVal,
    '==': mapAss,
    '+': mapSum,
    '*': mapProd,
    '?': mapCond,
    }
