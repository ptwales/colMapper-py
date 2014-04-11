import xlmp
import string
#import unittest

drng = list(range(0, 26))
xlExts = ['.xls', '.xlsx', '.ods']
testName = 'test'
Mt = [[string.uppercase[c] + str(r + 1) for r in drng] for c in drng]


if __name__ == '__main__':
    for x in xlExts:
        testbook = testName + x
        xlmp.byCol_not_byRow = True
        xlmp.__writeBook(Mt, testbook, testName)
        pass  # learn unittest
