class BaseContainer(object):

    def __init__(self):
        self._list = list()

    def clean_list(self):
        self._list = list()

    def get_list(self):
        return self._list

    def get_list_items_count(self):
        return len(self._list)
