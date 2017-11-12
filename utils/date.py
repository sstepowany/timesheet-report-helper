import datetime
import calendar


class Date(object):

    def __init__(self):
        self._first_month_day_format = "{}-{}-01"
        self._last_month_day_format = "{}-{}-{}"
        self._weekend_days_numbers = [5, 6]

    @staticmethod
    def get_current_year():
        return datetime.date.today().year

    @staticmethod
    def get_month_days_count(year, month):
        month_last_day = calendar.monthrange(year, month)[1]
        return month_last_day

    @staticmethod
    def get_month_name(month_number):
        return calendar.month_name[month_number]

    @staticmethod
    def get_current_date_time():
        return datetime.datetime.now()

    @staticmethod
    def parse_worklog_date_time(datetime_string):
        worklog_date_format = "%Y-%M-%d"
        worklog_date = datetime_string.split("T")[0]
        worklog_year = datetime.datetime.strptime(worklog_date, worklog_date_format).strftime("%Y")
        worklog_month = datetime.datetime.strptime(worklog_date, worklog_date_format).strftime("%M")
        worklog_day = datetime.datetime.strptime(worklog_date, worklog_date_format).strftime("%d")
        return worklog_year, worklog_month, worklog_day

    def get_month_range(self, year, month):
        month_last_day = self.get_month_days_count(year, month)
        first_month_day = self._first_month_day_format.format(year, month)
        last_month_day = self._last_month_day_format.format(year, month, month_last_day)
        return first_month_day, last_month_day

    def get_timesheet_year(self, timesheet_year):
        if timesheet_year == 'current':
            timesheet_year = self.get_current_year()
        return timesheet_year

    def is_day_of_weekend(self, year, month, day):
        weekday = datetime.datetime(int(year), int(month), int(day)).weekday()
        if weekday in self._weekend_days_numbers:
            is_day_of_weekend = True
        else:
            is_day_of_weekend = False
        return is_day_of_weekend
