colMapper-py
============

colMapper provides column mapping controls for Office Spreadsheet applications (MS Excel, LibreOffice-Calc, etc).
It maps columns from one spreadsheet to the another row by row.
_Number of rows in will always be the number of rows out_.
Later, I might add the ability to map between rows, but this is it for now.


## Dependencies

- Python [xlrd module](https://github.com/python-excel/xlrd)
- Python [xlwt module](https://github.com/python-excel/xlwt)

## Usage

Maps are defined with a `defaultdic(list)` object.
Each key is either a string representing a column name or an integer representing a column index, of the destination worksheet.
**Column indexes must be zero-offset** just as they are in `xlrd` and `xlwt`.

The data of each key can be a `list`, `int`, or `str`.
Any `int` is interperted as an column index of the source worksheet.
Any string will be interperted as a raw value.
A list will be evaluated as a function in Polish notation.

Consider this example from `demo_01.py`,

    demoCmd = defaultdict(list)
    demoCmd = {
        'A': 'Hello, World',
        'B': [Sum, 0, 1, 2],
        'C': [mapOper.mapAssert, 3, 4, 5],
         3 : [mapOper.mapProd, [mapOper.mapSum, 1, 0], 0, 1],
        'E': 5
        }
        
Column `'A'` will become 'Hello, World' in every cell.
Column `'E'` will become the column `'F'` of the source worksheet.
`'B'` will be the sum of columns `'A'`, `'B'`, and `'C'` from the source worksheet.
Column `'D'` (column index `3`) will be `('A'+'B')*'A'*'B'`.

`mapOper.py` intends to contain convienent functions to use.
_EG_ `mapOper.mapSum` filters null cells which would cause standard Python `Sum()` to raise an error.

Note there is an inconvience in the convention that _all strings are values and all integers are indexes_.
If you want destination column `'A'` to be a single increment of source column `'A'`. 
Then the command must be,

    'A': [Sum, 0, [int, '1']]  # Untested
        
as `[Sum, 0, '1']` will translate to `sheet.Cells(r,0).value + '1'` which will error if `sheet.Cells(r,0).value` is not a string.

To actually map the worksheets, use

    colMapper.interpColMap(demoCmd, fws, tws)
    
where `fws` is a `xlrd.open_workbook.sheet` object and `tws` is and `xlwt.Workbook.sheet` object.

##TODO

- Add more functions.
  + re function
- row expansion or compression
- Maybe
  + Create `mapInterpertor`
    - a factory that accepts a string and interperts the apropriate function
    
