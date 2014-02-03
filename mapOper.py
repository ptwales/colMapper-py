class ImapOper(object):

    def __init__(self):

    def Evaluate(self, valueList):
        print ("Evaluate not Initialized")

    def getToken(self):
        return self.token
# operators do not touch the worksheets they are printing to
# They are functional objects that are to be stacked, queued, listed, etc
#    def __stdList__(self, mapCmd):
#        for i in range(0, len(mapCmd)):
#            for row in range(1, sheet.UsedRange.Rows.Count):
#                values[i][row] = sheet.cells(row, mapCmd[i]).value
#        return values


class mapIs(ImapOper):

    def __init__(self):
        self.token = "="

    def Evaluate(self, valueList):
        return values[0]


class mapSum(ImapOper):
    
    def __init__(self):
        self.token = "+"

    def Evaluate(self, valueList):
        return sum(valueList)


class mapAss(ImapOper):
    
    def __init__(self):
        self.token = "=="
        self.ASSERTFAIL = "~ASSERT_FAIL!"
    
    def Evaluate(self, valueList):
        # Later Optionally remove Null Values from the list
        # self.__stripNull__(valueList)
        if len(valueList) == 1:
            return valueList[0]
        for i in range(0, len(valueList) - 1):
            if valueList[i] != valueList[i+1]:
                return self.ASSERTFAIL
        return valueList[0]
    
    def __stripNULL__(self, valueList):
        # remove Null values from the list


class mapCond(ImapOper):
    
    def __init__(self):
        self.token = "?"
    
    def Evaluate(self, valueList):
        # valueList = (BOOL, TRUE_VALUE, FALSE_VALUE)
        assert len(valueList) = 3
        return valueList[1] if valueList[0] else valueList[2]
