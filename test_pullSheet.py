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


if __name__ == '__main__':
    makeTestBook('test.ods', 'test')
    book = xlrd.open_workbook('test.ods')
    sheet = book.sheet_by_index(0)
    M = xlmp.__pullSheet(sheet)
    print(M)
