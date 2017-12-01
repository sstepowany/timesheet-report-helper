class DataConverter(object):

    def __init__(self):
        self._ranges_separator = ","
        self._range_separator = "-"
        self._empty_days_list_maker = 'None'

    def split_configuration_special_days_to_list(self, list_configuration):
        days_list = list()
        if list_configuration != self._empty_days_list_maker:
            days_ranges_and_days_list = str(list_configuration).split(self._ranges_separator)
            for days_range_or_day in days_ranges_and_days_list:
                if self._range_separator in days_range_or_day:
                    days_range = days_range_or_day.split(self._range_separator)
                    for day in range(int(days_range[0]), int(days_range[1]) +1):
                        days_list.append(str(day))
                else:
                    days_list.append(days_range_or_day)
        return days_list
