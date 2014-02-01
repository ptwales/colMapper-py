class ImapOper(object):
	def Evaluate(mapCmd[]):
		print "Evaluate not Initialized"
		return
	def getToken(self):
		return self.token
	def __stdList__(mapCmd[]):
		for i in range(0, len(mapCmd))
			for row in range(1, sheet.UsedRange.Rows.Count)
				values[i][row] = sheet.cells(row,mapCmd[i]).value
		return values
	pass

class mapIs(ImapOper):
	def Evaluate(mapCmd[]):
		return mapCmd[0]
	pass

class mapSum(ImapOper):
	def Evaluate(mapCmd[]):
		row_i = 2
		for row in self.__stdList__(mapCmd)
			value = 0
			for col in row
				value = col + value
			sheet.cells(row_i).value = value
		return
	pass
