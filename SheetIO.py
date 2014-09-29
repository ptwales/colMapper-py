import xlrd
import xlwt

class ISheetIO(object):
    """ Abstract base class for SheetIO
    """
    def read_sheet(self):
        raise NotImplementedError("read_sheet not implemented")
    
    def write_sheet(self):
        raise NotImplementedError("write_sheet not implemented")
    
    
    
class SheetIO(ISheetIO):
  
    def __init__(self, read_by_row=True):
        self.by_row = read_by_row
      
    def _factory(self, sheet):
        sheet_type = type(sheet)
        if sheet_type is xlrd.sheet or sheet_type is xlwt.WorkSheet:
            return ExcelSheet(read_by_row)
        # elif CSV
        # elif OpenXML
        else:
            raise TypeError("Unknown Sheet type or not a sheet")
    
    def read_sheet(self, sheet, **kwargs):
        data = self._factory(sheet).read_sheet(sheet, **kwargs)
        return self.orient_data(data)
    
    def write_sheet(self, data, sheet, **kwargs):
        oriented_data = self._orient_data(data)
        self._factory(sheet).write_sheet(sheet, oriented_data, **kwargs)
    
    def _orient_data(self, data):
        if self.by_row:
            return data
        else:
            return self._transpose(data) 
  
    def _transpose(self, matrix):
        return zip(*matrix)



class ExcelSheet(ISheetIO):

    def read_sheet(self, sheet,
            col_start=0, row_start=0, col_cnt=-1, row_cnt=-1):
        
        # TODO: Raise an error instead. Research which error to raise
        assert not sheet.ragged_rows
        
        if row_cnt < 0:
            row_cnt += sheet.nrows + 1
        if col_cnt < 0:
            col_cnt += sheet.ncols + 1
            
        return [sheet.row_values(r, col_start, row_start) 
                for r in xrange(row_start, row_cnt + 1)]

    def write_sheet(self, data, sheet, col_start=0, row_start=0):
        for r, row, in enumerate(data):
            for c, el in enumerate(row):
                sheet.write(row_start + r, col_start + c, el)



class CsvSheet(ISheetIO):

    def read_sheet(self, read_path, 
            col_start=0, row_start=0, col_cnt=-1, row_cnt=-1):
        with open(read_path, 'rb') as csv_file:
            dialect = csv.Sniffer().sniff(csv_file.readline())
            csv_file.seek(0)
            csv_reader = csv.reader(csvfile, dialect))
            data = [row[col_start:col_cnt] for row in 
                    csv_reader][row_start:row_cnt]
        return data

    def write_sheet(self, data, write_path, **kwargs):
        with open(write_path, 'wb') as csv_file:
            csv_writer = csv.writer(csv_file, **kwargs)
            csv_writer.writerows(data)



# class OpenXMLSheet(ISheetIO):
