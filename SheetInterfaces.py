import xlrd
import xlwt

class ISheet(object):
  
  def __init__(self, read_by_row=True):
      self.by_row = read_by_row
  
  def orient_data(self, data_matrix):
      if self.by_row:
          return data_matrix
      else:
          return self._transpose(data_matrix) 
  
  def _transpose(self, data_matrix):
      return zip(*data_matrix)



class ExcelSheet(ISheet):

    def read_sheet(self, sheet, col_start=0, row_start=0, col_cnt=-1, row_cnt=-1):
        
        # TODO: Raise an error instead. Research which error to raise
        assert not sheet.ragged_rows
        if row_cnt < 0:
            row_cnt += sheet.nrows + 1
            
        # xlrd.sheet.row_values uses row[c0:cf]
        # therefore cf=-1 excludes last.
        if col_cnt < 0:
            col_cnt += sheet.ncols + 1
            
        data_matrix = [sheet.row_values(r, col_start, row_start) 
                       for r in xrange(row_start, row_cnt + 1)]
        return super(ExcelSheet, self).orient_data(data_matrix)

    def write_sheet(self, data_matrix, sheet, col_start=0, row_start=0):
    
        oriented_matrix = super(ExcelSheet, self).orient_data(data_matrix)
        
        for r, row, in enumerate(oriented_matrix):
            for c, el in enumerate(row):
                sheet.write(row_start + r, col_start + c, el)

# class CSVSheet(ISheet):
# class OpenXMLSheet(ISheet):
