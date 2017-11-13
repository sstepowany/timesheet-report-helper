from timesheets.worksheets.worksheetObjects.worksheet_column import WorksheetColumn


class WorksheetColumnsRange(object):

    def __init__(self, first_column_index, last_column_index, columns_width):
        self._first_column = WorksheetColumn(first_column_index, columns_width)
        self._last_column = WorksheetColumn(last_column_index, columns_width)

    def get_first_column(self):
        return self._first_column

    def get_range_columns_coordinates(self):
        first_column_index = self._first_column.get_column_index()
        last_column_index = self._last_column.get_column_index()
        return first_column_index, last_column_index
