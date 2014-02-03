class ImapOper(object):

    def __init__(self):

    def Evaluate(self, mapCmd):
        print ("Evaluate not Initialized")

    def getToken(self):
        return self.token

    def __stdList__(self, mapCmd):
        for i in range(0, len(mapCmd)):
            for row in range(1, sheet.UsedRange.Rows.Count):
                values[i][row] = sheet.cells(row, mapCmd[i]).value
        return values


class mapIs(ImapOper):

    def Evaluate(self, mapCmd):
        return mapCmd[0]


class mapSum(ImapOper):

    def Evaluate(self, mapCmd):
        row_i = 2
        for row in self.__stdList__(mapCmd):
            value = 0
            for col in row
                value = col + value
                sheet.cells(row_i).value = value
