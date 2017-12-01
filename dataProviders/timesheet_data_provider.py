from utils.date import Date
from rest.jira_rest import JiraRest
from utils.employee_jira_data_parser import EmployeeJiraDataParser
from utils.logger import TimesheetLogger


class TimesheetData(object):

    def __init__(self, jira_rest_configuration, timesheet_date):
        self._jira_rest = JiraRest(jira_rest_configuration)
        self._timesheet_date = timesheet_date
        self._employee_jira_data_parser = EmployeeJiraDataParser()

    def get_timesheet_data_for_employee(self, employee_jira_name):
        TimesheetLogger().log_on_console("Getting employee: {}, timesheet data.".format(employee_jira_name))
        timesheet_year = Date().get_timesheet_year(self._timesheet_date['year'])
        params = self._jira_rest.prepare_search_parameters_for_timesheet_data(
            employee_jira_name, timesheet_year, self._timesheet_date['month'])
        response = self._jira_rest.search_for_data(params)
        return response.json()

    def get_timesheet_worklog_data(self, employee_timesheet_issues, employee_jira_name):
        TimesheetLogger().log_on_console("Getting employee: {}, worklog data.".format(employee_jira_name))
        employee_worklog_data = dict()
        for issue in employee_timesheet_issues:
            issue_key = issue['key']
            issue_id = issue['issue_id']
            response = self._jira_rest.get_issue_worklog(issue_id)
            employee_worklog_data[issue_key] = list()
            issue_worklog_data = self._employee_jira_data_parser.parse_employe_worklog_data_2(
                employee_worklog_data[issue_key], employee_jira_name, response.json())
            employee_worklog_data[issue_key] = issue_worklog_data
        return employee_worklog_data

    def get_employee_display_name(self, employee_jira_name):
        TimesheetLogger().log_on_console("Getting employee: {}, display name.".format(employee_jira_name))
        params = self._jira_rest.perapare_params_for_user_rest(employee_jira_name)
        response = self._jira_rest.get_user_data(params)
        employee_display_name = response.json()['displayName']
        return employee_display_name
