class WorksheetCell(object):

    def __init__(self, row_index, column_index, cell_content, cell_format):
        self._row_index = row_index
        self._column_index = column_index
        self._cell_content = cell_content
        self._format = cell_format

    def get_cell_coordinates(self):
        return self._row_index, self._column_index

    def get_cell_content(self):
        return self._cell_content

    def get_cell_format(self):
        return self._format
