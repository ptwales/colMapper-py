colMapper-py
============

Python column mapping controls for Office Spreadsheet applications (MS Excel, LibreOffice-Calc, etc).
It maps columns from one spreadsheet to the another by row.  _Each row is independent of other rows._

Maps: 
  - one column to one column
  - one column to multiple columns
  - Multiple columns to one column with specified handling
    + Assert that each column have the same value, return that value.
    + Return the sum of each value of the columns
    + Return the product of each value of the columns
  - More?

##TODO

- Add more functions.
- Functions that accept multiple columns should optionally, but default, ignore null values.
- Maybe
  + Create `mapInterpertor`
    - is a factory that accepts a string and interperts the apropriate function

## Dependencies

- Python [xlrd module](https://github.com/python-excel/xlrd)
- Python [xlwt module](https://github.com/python-excel/xlwt)
