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

Maps are defined with a `defaultdict(list)` object.
Each key is either a string representing a column name or an integer representing a column index, of the destination worksheet.
**Column indexes must be zero-offset** just as they are in xlrd and xlwt.

The data of each key can be a list, integer, or string.
Any integer is interperted as an column index of the source worksheet.
Any string will be interperted as a raw value.
A list will be evaluated as a function in Polish notation.

Consider this example from demo_01.py,

    demoCmd = defaultdict(list)
    demoCmd = {
        'A': 'Hello, World',
        'B': [Sum, 0, 1, 2],
        'C': [mapOper.mapAssert, 3, 4, 5],
         3 : [mapOper.mapProd, [mapOper.mapSum, 1, 0], 0, 1],
        'E': 5
        }
        
Column _A_ will become 'Hello, World' in every cell.
Column _E_ will become the column _F_ of the source worksheet.
_B_ will be the sum of columns _A_, _B_, and _C_ from the source worksheet.
Column _D_ (column index 3) will be `('A'+'B')*'A'*'B'`.

`mapOper.py` intends to contain convienent functions to use.
_EG_ `mapOper.mapSum` filters null cells which would cause standard Python `Sum()` to raise an error.

Note there is an inconvience in the convention that _all strings are values and all integers are indexes_.
If you want destination column _A_ to be a single increment of source column _A_. 
Then the command must be,

    'A': [Sum, 0, [int, '1']]  # Untested
        
as `[Sum, 0, '1']` will translate to `sheet.Cells(r,0).value + '1'` which will error if `sheet.Cells(r,0).value` is not a string.

To actually map the worksheets, use

    colMapper.interpColMap(demoCmd, fws, tws)
    
where `fws` is a `xlrd.open_workbook.sheet` object and `tws` is and `xlwt.Workbook.sheet` object.

## License

  GPLv3

## To Do

- Add more functions.
  + re function
- Allow row maping under the same priciple of colMapping
  + Will need to provide dynamic creation or looping of map commands
- Create a map interperter
  + a factory that accepts a string and interperts the apropriate function
    
