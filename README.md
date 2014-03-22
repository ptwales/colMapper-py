colMapper-py
============

Whoa! Seriously overhauling the internals.
Get back to you on how it works later

colMapper provides column mapping controls for Office Spreadsheet applications (MS Excel, LibreOffice-Calc, etc).
It maps columns from one spreadsheet to the another row by row.
_Number of rows in will always be the number of rows out_.
Later, I might add the ability to map between rows, but this is it for now.


## Dependencies

- Python [xlrd module](https://github.com/python-excel/xlrd)
- Python [xlwt module](https://github.com/python-excel/xlwt)
- Python 2.3 - 2.7 (As needed by above)

## Usage

Maps are defined with a `defaultdict(tuple)` object.
Recently changed. 
See `demo/demo_1.py` for details.
Will update this section later.

## License

  LGPLv3

## To Do

- Add more functions.
  + re function
- Allow row maping under the same priciple of colMapping
  + Will need to provide dynamic creation or looping of map commands
- Create a map interperter
  + a factory that accepts a string and interperts the apropriate function
    
