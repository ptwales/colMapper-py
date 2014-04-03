import xlwt
import xlrd
from xlmp import xlmp
import string


def makeTestBook(bookName, sheetName):
    book = xlwt.Workbook()
    sheet = book.add_sheet(sheetName)
    for r in range(0, 26):
        for c in range(0, 26):
            sheet.write(r, c, string.uppercase[c] + str(r + 1))
    book.save(bookName)

xlExts = ['.xls', '.xlsx', '.ods']
testName = 'test'
global map_byCol_not_byRow

if __name__ == '__main__':
    for x in xlExts:
        xlmp.map_byCol_not_byRow = True
        makeTestBook(testName + x, testName)
        book = xlrd.open_workbook(testName + x)
        sheet = book.sheet_by_index(0)

        # test xlmp.pullsheet
        M = xlmp.pullSheet(sheet)
        print('Basic Pull' + '-' * 50)
        print(M)

        # test xlmp.writeSheet
        cpyBk = xlwt.Workbook()
        cpySh = cpyBk.add_sheet(testName)
        xlmp.writeSheet(cpySh, M)
        cpyBk.save(testName + x)

        # test xlmp.map_byCol_not_byRow
        xlmp.map_byCol_not_byRow = False
        M = xlmp.pullSheet(sheet)
        print('Transposed Pull' + '-' * 50)
        print(M)
        tnpsBk = xlwt.Workbook()
        tnpsSh = tnpsBk.add_sheet(testName)
        xlmp.writeSheet(tnpsSh, M)
        tnpsBk.save('Trns' + testName + x)

