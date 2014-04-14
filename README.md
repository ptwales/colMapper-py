xlmp
============

xlmp provides mapping controls for Office Spreadsheet applications (MS Excel, LibreOffice-Calc, etc).
It can do row mapping, column mapping, and sub-mapping.
Mapping is based on a similar procedures as matrix multiplication meaning xlmp maps by either columns or rows at a time.  
This means that the number of rows or columns in equal the number of rows or columns out if you map by columns or rows, respectively.
With the exception of sub-mapping, where xlmp partitions the data set along one axis and then maps each sub-set along the other axis.

## Usage

Mapping is done by defining an `xlmp.mpCmd`,

```
>>> myMapCommand = xlmp.mpCmd(zeroOffSet=0, {'A': 0})

```
This particular command would just map column `A` to column `A` and passing it to the xlmp function,
```
>>> xlmp.xlmp(myMapCommand, fBook="fromBook.xls", tBook="output.xls")
```

## Methodology

xlmp treats ranges of spread-sheets as matricies without regarding the data type of any cell.
All values of the specified origin range are read as rectangular list of lists through an interface to `xlrd` and `xlwt`.
xlmp performs the matrix operation,

```
$ \mathbb{B} = \mathbb{M} \cdot \mathbb{G} $
```
Where **M** is the original dataset and **B** is the desired output dataset.
**G** is a column vector of functions which accept a vector and returns a scalar.
This is where xlmp differs from normal _(def?)_ matrix multiplication.
Instead of element `b_{i,j}` equaling the dot product of row vector `\vec{m}_i` and `\vec{g}_j`, elements of **B** are defined as,

```
$ b_{i,j} = g_j \left(\vec{m}_i\right) $
```

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


    
