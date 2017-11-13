from timesheets.worksheets.worksheetObjects.worksheet_cell import WorksheetCell


class WorksheetCellsRange(object):

    def __init__(self, first_cell_row_index, first_cell_column_index, second_cell_row_index, second_cell_column_index,
                 range_content, range_format):
        self._range_first_cell = \
            WorksheetCell(first_cell_row_index, first_cell_column_index, range_content, range_format)
        self._range_second_cell = \
            WorksheetCell(second_cell_row_index, second_cell_column_index, range_content, range_format)

    def get_range_second_cell(self):
        return self._range_second_cell

    def get_range_cells_coordinates(self):
        first_cell_coordinates = self._range_first_cell.get_cell_coordinates()
        second_cell_coordinates = self._range_second_cell.get_cell_coordinates()
        first_cell_row, first_cell_column = first_cell_coordinates
        second_cell_row, second_cell_column = second_cell_coordinates
        return first_cell_row, first_cell_column, second_cell_row, second_cell_column
