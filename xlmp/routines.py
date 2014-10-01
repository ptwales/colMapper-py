from itertools import groupby


# performs [M].[F] = B
# where F_j(M_i) = B[i][j]
# default is by cols of F and rows of M
# M must be transposed beforehand for other method
def _command_operate(map_command, data_matrix):
    
    x_range = range(max(map_command.keys()) + 1)
    y_range = range(len(data_matrix[0]))
    
    # preallocate because some columns could be skipped
    resultant_matrix = [[None for i in x_range] for j in y_range]
    assert len(x_range) == len(data_matrix)
    
    for i, row in enumerate(data_matrix):
        for k in map_command.keys():
            resultant_matrix[i][k] = map_command[k](row)
            
    return resultant_matrix

def group_by_ids(data_matrix, id_indexes):

    def id_func(x):
        return [x[i] for i in id_indexes]

    data_matrix.sort(key=id_indexes)
    return [list(g) for k, g in groupby(data_matrix, id_func)]

# Wrapper functions

def line_mapping(line_map, in_sheet, out_sheet, map_by_row=True, 
        read_kwargs={}, write_kwargs={}):
    
    sheet_io = SheetIO(map_by_row)
    data_matrix = sheet_io.read_sheet(in_sheet, **read_kwargs)
    result_matrix = _command_operate(line_map, data_matrix)
    sheet_io.write_sheet(out_sheet, **write_kwargs)


def block_mapping(block_map, in_sheet, out_sheet, grp_func, grp_func_kwargs={},
        grp_by_col=True, read_kwargs={}, write_kwargs={}):
    
    sheet_io = SheetIO(by_row)
    data_matrix = sheet_io.read_sheet(in_sheet, **read_kwargs)
    
    block_matrix = grp_func(data_matrix, **grp_func_kwargs)
    # map each block then merge the mapped_blocks
    mapped_matrix = zip([_command_operate(block_map, zip(*block))
                        for block in block_matrix])
    
    sheet_io.write_sheet(mapped_matrix, **write_kwargs)

