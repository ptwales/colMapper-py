colMapper-py
============

Python column mapping controls for Office Spreadsheet applications (MS Excel, LibreOffice-Calc, etc)

Maps columns from one spreadsheet to the another
Maps: 
  - one column to one column
  - one column to multiple columns
  - one value to all columns
  - Multiple columns to one column with specified handling
    + Assert that both columns are the same
    + Sum values in both columns
    + ???

###TODO

- turn `mapOpers` into pure functions
- add more functions
- `mapAss` needs to optionally remove Null Values
- add queue to `mapper`
- add flow control to `mapper`
- Maybe
  + Create `mapInterpertor`
    - is a factory that accepts a string and interperts the apropriate function
