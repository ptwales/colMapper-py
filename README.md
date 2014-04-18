xlmp
====

xlmp provides mapping controls for Office Spreadsheet applications (MS Excel, LibreOffice-Calc, etc).
It can do row mapping, column mapping, and sub-mapping.
Mapping is based on a similar procedures as matrix multiplication meaning xlmp maps by either columns or rows at a time.  
This means that the number of rows or columns in equal the number of rows or columns out if you map by columns or rows, respectively.
With the exception of sub-mapping, where xlmp partitions the data set along one axis and then maps each sub-set along the other axis.

## Usage

Mapping is done by defining an _xlmp.mpCmd_ object,

```python
>>> import xlmp
>>> mapCommand = xlmp.mpCmd({'A': 0})
```

and passing it to the xlmp function,

```python
>>> xlmp.xlmp(mapCommand, fBook="fromBook.xls", tBook="output.xls")
```

would result in a file _output.xls_ which would only contain to column 'A' of _fromBook.xls_ in column 'A'.

### xlmp.mpCmd

#### Indexes

A simple map that reorders the a sheet would look like,

```python
>>> mapCommand = xlmp.mpCmd({'A': 2, 2: 1, 'C': 4, 'e': 5}, offset=1)
```

Will map column 'A' &rarr; 'B', 'B' &rarr; 'A', 'C' &rarr; 'D', and 'E' &rarr; 'E'.
The row count does not change and no values copied are altered.
Note that column names are case insensitive and can be used for origins but not _yet_ for destinations.
Furthermore, the optional `offset` parameter can be set to control the offset of indexes but not names.
'A' &#x2194; 'A' regardless of `offset`.

To reorder rows instead of columns, set the optional arguement `byCol=False` when performing the map,

```python
>>> xlmp.xlmp(mapCommand, byCol=False, fBook="fromBook.xls", tBook="output.xls")
```

A map using column names can still be used for row mapping; 'A' &rarr; row 1, 'AA' &rarr; row 27, etc.

#### Functions

To map a new value as a function of other values, pass a function and a list of indexes instead of just an index.
The list of indexes will specify values of which cells on that row or column to pass to that function.

```python
>>> mapCommand = xlmp.mpCmd({'A': (Sum, [0, 3, 4])})
```
will map 'A' = Sum('A','D','E'), ergo 'A1' = Sum('A1', 'D1', 'E1'), if mapping by columns.

## Methodology

xlmp treats ranges of spread-sheets as matricies without regarding the data type of any cell.
All values of the specified origin range are read as rectangular list of lists through an interface to xlrd and xlwt.
xlmp performs the matrix operation,

<span class="align-center">
  <a href="http://www.codecogs.com/eqnedit.php?latex=\textbf{B}=\textbf{M}\cdot\vec{G}"target="_blank">
    <img src="http://latex.codecogs.com/gif.latex?\textbf{B}=\textbf{M}\cdot\vec{G}"title="\textbf{B}=\textbf{M}\cdot\vec{G}"/>
  </a>
</span>

to the dataset, where **M** is the original dataset and **B** is the desired output dataset.
The output is written to a file through the same interface.

_G_<span style="position: relative;bottom: 1.3ex;letter-spacing: -1.2ex;right: 1.2ex">_&#x20d7;_</span>
is a column vector of functions which accept a vector and returns a scalar.
This is where xlmp differs from normal _(def?)_ matrix multiplication.
Instead of element _b<sub>i,j</sub>_ equaling the dot product of row vector _m<sub>i</sub>_ and _g<sub>j</sub>_, elements of **B** are defined as,

<span class="align-center">
  <a href="http://www.codecogs.com/eqnedit.php?latex=b_{i,j}=g_j\left(m_i\right)" target="_blank">
    <img src="http://latex.codecogs.com/gif.latex?b_{i,j}=g_j\left(m_i\right)" title=" b_{i,j}=g_j\left(m_i\right)"/>
  </a>
</span>

<!--
### Map Commands

In practice, not all of elements of _m<sub>i</sub>_ would be needed by _g<sub>j</sub>_.
Constraining the user to map with only function that accept iterables is unnecessary.
-->

### Transposition

Note the constraint that **B** must have the row count as **M**.
This can be circumvented by performing the operation,

<!-- LMAO if this works -->
```LaTeX
\textbf{B}=\vec{G}\cdot\textbf{M}
```

constraining the column count instead.
In the code, xlmp does that operation as,

<!-- LMAO if this works -->
```LaTeX
\textbf{B}^{T}=\textbf{M}^{T}\vec{G}^{T}
```

by reading M and writing B transposed through the xlrd-xlwt interface.
_&#x20d7;G_ does not need to be transposed in practice as it is just a dictionary of functions.

### Sub Mapping

## Dependencies

- Python [xlrd module](https://github.com/python-excel/xlrd)
- Python [xlwt module](https://github.com/python-excel/xlwt)
- Python 2.3 - 2.7 (As needed by above)

## License
 
 Copywrite (C) 2014 Philip Wales

 This software is licensed under the GNU Lesser General Public License (LGPL), version 3 ("the License").
 See the License for details about distribution rights, and the specific rights regarding derivate works.
 You may obtain a copy of the License at:
 
 http://www.gnu.org/licenses/licenses.html

    
