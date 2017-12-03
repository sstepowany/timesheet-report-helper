from timesheets.timesheetFormats import TimesheetFormats
from timesheets.worksheets.employee_worksheet import EmployeeWorksheet
from timesheets.worksheets.statistics_worksheet import StatisticsWorksheet
from timesheets.worksheets.summary_wokrsheet import SummaryWorksheet
from utils.logger import TimesheetLogger


class Timesheet(TimesheetFormats):

    def __init__(self, workbook):
        self._workbook = workbook
        self._statistics_worksheet_name = "Statistics"
        self._summary_worksheet_name = "Summary"
        super(Timesheet, self).__init__(workbook)

        self._employee_worksheet = EmployeeWorksheet(self._formats_dict)
        self._statistics_worksheet = StatisticsWorksheet(self._formats_dict)
        self._summary_worksheet = SummaryWorksheet(self._formats_dict)

    def _add_employee_worksheet(self, employee, timesheet_date, man_days_hours):
        employee_jira_name = employee.get_emplyee_jira_name()
        TimesheetLogger().log_on_console("*" * 50)
        TimesheetLogger().log_on_console("Creating worksheet for employee: {}.".format(employee_jira_name))
        employee_name = employee.get_employee_name()
        employee_timesheet_issues = employee.get_employee_timesheet_issues()
        employee_worksheet = self._workbook.add_worksheet(employee_name)
        employee_worklog_data = employee.get_employee_timesheet_worklog_data()
        self._employee_worksheet.prepare_worksheet_for_employee(employee_worksheet, employee_timesheet_issues,
                                                                employee_worklog_data, timesheet_date, man_days_hours)

    def add_worksheets_for_employees(self, teams_list, timesheet_date, man_days_hours):
        for team in teams_list:
            for employee in team.get_employees_list():
                self._add_employee_worksheet(employee, timesheet_date, man_days_hours)

    def add_statistics_worksheet(self, teams_list, timesheet_date, man_days_hours):
        TimesheetLogger().log_on_console("*" * 50)
        TimesheetLogger().log_on_console("Creating statistics worksheet.")
        statistics_worksheet = self._workbook.add_worksheet(self._statistics_worksheet_name)
        self._statistics_worksheet.prepare_statistics_worksheet(statistics_worksheet, teams_list, timesheet_date,
                                                                man_days_hours)

    def add_summary_worksheet(self, teams_list, timesheet_date, national_holidays_days_list, po_number, man_days_costs):
        TimesheetLogger().log_on_console("*" * 50)
        TimesheetLogger().log_on_console("Creating summary worksheet.")
        summary_worksheet = self._workbook.add_worksheet(self._summary_worksheet_name)
        self._summary_worksheet.prepare_summary_worksheet(summary_worksheet, teams_list, timesheet_date,
                                                          national_holidays_days_list, po_number, man_days_costs)

    def save_timesheet(self):
        self._workbook.close()
