xlmp
====

xlmp provides mapping controls for Office Spreadsheet applications (MS Excel, LibreOffice-Calc, etc).
It can do row mapping, column mapping, and sub-mapping.
Mapping is based on a similar procedures as matrix multiplication meaning xlmp maps by either columns or rows at a time.
This means that the number of rows or columns in equal the number of rows or columns out if you map by columns or rows, respectively.
With the exception of sub-mapping, where xlmp partitions the data set along one axis and then maps each sub-set along the other axis.


### xlmp.mpCmd

#### Indexes

A simple map that reorders the a sheet would look like,

    mapCommand = xlmp.mpCmd({'A': 2, 2: 1, 'C': 4, 'e': 5}, offset=1)

Will map column 'A' &rarr; 'B', 'B' &rarr; 'A', 'C' &rarr; 'D', and 'E' &rarr; 'E'.
The row count does not change and no values copied are altered.
Note that column names are case insensitive and can be used for origins but not _yet_ for destinations.
Furthermore, the optional _offset_ parameter can be set to control the offset of indexes but not names.
'A' &#x2194; 'A' regardless of _offset_.

#### Functions

To map a new value as a function of other values, pass a function and a list of indexes instead of just an index.
The list of indexes will specify values of which cells on that row or column to pass to that function.

    mapCommand = xlmp.mpCmd({'A': (Sum, [0, 3, 4])})

will map 'A' = Sum('A','D','E'), ergo 'A1' = Sum('A1', 'D1', 'E1'), if mapping by columns.


Dependencies
------------

- Python [xlrd module](https://github.com/python-excel/xlrd)
- Python [xlwt module](https://github.com/python-excel/xlwt)
- Python 2.3 - 2.7 (As needed by above)

License
-------
 
> Copyright (C) 2014 Philip Wales
> 
>  This software is licensed under the GNU Lesser General Public License (LGPL), version 3 ("the License").
>  See the License for details about distribution rights, and the specific rights regarding derivate works.
>  You may obtain a copy of the License at:
>
> <http://www.gnu.org/licenses/licenses.html>

