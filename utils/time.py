class Time(object):

    def __init__(self):
        self._secnods_in_minute = 60
        self._miuntes_in_hour = 60
        self._hours_time_format = "{}.{}"

    def parse_seconds_to_hours(self, seconds):
        minutes = seconds / self._secnods_in_minute
        hours = minutes / self._miuntes_in_hour
        minutes_left = minutes % self._secnods_in_minute
        hour_fraction = 100 * minutes_left / self._miuntes_in_hour
        return self._hours_time_format.format(hours, hour_fraction)
