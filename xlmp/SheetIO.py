import xlrd
import xlwt
import csv


def factory_create_sheet(sheet):
    """ Factory Method, Decide which concrete
    class to return based on the sheet type.
    """
    factory_registry = {
        xlrd.sheet: ExcelSheetIO(),
        xlwt.WorkSheet: ExcelSheetIO(),
        csv.reader: CsvSheetIO(),
        csv.writer: CsvSheetIO()
    }
    try:
        return factory_registry[type(sheet)]
    except(KeyError):
        raise TypeError("Unknown Sheet type or not a sheet")


def transpose(matrix):
    return zip(*matrix)


class UniversalSheetIO(object):
    """ Universal Reader that implements factory method and
    matrix transposition.
    """
    def __init__(self, read_by_row=True):
        self.by_row = read_by_row
        
    def read_sheet(self, sheet, **kwargs):
        data = factory_create_sheet(sheet).read_sheet(sheet, **kwargs)
        return self.orient_data(data)
    
    def write_sheet(self, data, sheet, **kwargs):
        oriented_data = self._orient_data(data)
        factory_create_sheet(sheet).write_sheet(sheet, oriented_data, **kwargs)
    
    def _orient_data(self, data):
        if self.by_row:
            return data
        else:
            return transpose(data)



class ConcreteSheetIO(object):
    """ Abstract base class for SheetIO
    """
    def read_sheet(self):
        raise NotImplementedError("read_sheet not implemented")
    
    def write_sheet(self):
        raise NotImplementedError("write_sheet not implemented")



class ExcelSheetIO(ConcreteSheetIO):
    """ Concrete implementation of Excel Sheets using xlrd and xlwt
    """
    def read_sheet(self, sheet,
            col_start=0, row_start=0, col_cnt=-1, row_cnt=-1):
        # TODO: Raise an error instead. Research which error to raise
        assert not sheet.ragged_rows
        # Allow wrapping
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



class CsvSheetIO(ConcreteSheetIO):
    """ Concreted implementation using csv module
    """
    def read_sheet(self, csv_reader, 
            col_start=0, row_start=0, col_cnt=-1, row_cnt=-1):
        return [row[col_start:col_cnt] for row in sv_reader][row_start:row_cnt]

    def write_sheet(self, data, csv_writer):
        csv_writer.writerows(data)



# class OpenXMLSheet(ISheetIO):
