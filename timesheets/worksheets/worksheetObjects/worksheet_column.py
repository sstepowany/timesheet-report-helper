class WorksheetColumn(object):

    def __init__(self, column_index, column_width):
        self._column_index = column_index
        self._column_width = column_width

    def get_column_index(self):
        return self._column_index

    def get_column_width(self):
        return self._column_width
