from utils.date import Date


class NamesGenerator(object):

    def __init__(self, timesheet_data):
        self._timesheet_data = timesheet_data
        self._report_name_template = "{} - timesheet report - {} - {}.xlsx"
        self._date_format_template = "{}.{}.{}"
        self._timesheets_reports_set_folder_name_template = "{} - {}"
        self._date = Date()

    def _prepare_generated_name_basic_data(self):
        timesheet_year = self._date.get_timesheet_year(self._timesheet_data['year'])
        timesheet_month = self._timesheet_data['month']
        month_length = self._date.get_month_days_count(timesheet_year, timesheet_month)
        from_date = self._date_format_template.format(timesheet_year, timesheet_month, 1)
        to_date = self._date_format_template.format(timesheet_year, timesheet_month, month_length)
        return from_date, to_date

    def generate_report_name(self, timesheet_name_prefix):
        from_date, to_date = self._prepare_generated_name_basic_data()
        report_name = self._report_name_template.format(timesheet_name_prefix, from_date, to_date)
        return report_name

    def generate_timesheets_set_folder_name(self):
        from_date, to_date = self._prepare_generated_name_basic_data()
        timesheets_reports_set_folder_name = self._timesheets_reports_set_folder_name_template.\
            format(from_date, to_date)
        return timesheets_reports_set_folder_name
