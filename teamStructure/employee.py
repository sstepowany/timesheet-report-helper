from utils.employee_jira_data_parser import EmployeeJiraDataParser
from utils.data_converter import DataConverter


class Employee(object):

    def __init__(self, employee_jira_name, employee_configuration_data, timesheet_data_object):
        self._employee_jira_name = employee_jira_name
        self._timesheet_data = timesheet_data_object
        self._employee_configuration_data = employee_configuration_data
        self._trainings_and_other_projects_key = 'trainings/other_projects'
        self._holidays_and_sick_leaves_key = 'holidays/sick_leaves'

        self._employee_name = self._timesheet_data.get_employee_display_name(self._employee_jira_name)
        self._employee_timesheet_data = self._timesheet_data.get_timesheet_data_for_employee(self._employee_jira_name)
        self._employee_timesheet_issues = EmployeeJiraDataParser().get_jira_issues(self._employee_timesheet_data)
        self._employee_timesheet_worklog_data = self._timesheet_data.get_timesheet_worklog_data(
            self._employee_timesheet_issues, self._employee_jira_name)

    def get_employee_name(self):
        return self._employee_name

    def get_emplyee_jira_name(self):
        return self._employee_jira_name

    def get_employee_configuration_holidays_and_sick_leaves_days_list(self):
        employee_holidays_and_sick_leaves_days_list = \
            DataConverter().split_configuration_special_days_to_list(
                self._employee_configuration_data[self._holidays_and_sick_leaves_key])
        return employee_holidays_and_sick_leaves_days_list

    def get_employee_configuration_trainings_and_other_projects_days_list(self):
        employee_trainings_and_other_projects_days_list = \
            DataConverter().split_configuration_special_days_to_list(
                self._employee_configuration_data[self._trainings_and_other_projects_key])
        return employee_trainings_and_other_projects_days_list

    def get_employee_configuration_level(self):
        return self._employee_configuration_data['level']

    def get_employee_timesheet_issues(self):
        return self._employee_timesheet_issues

    def get_employee_timesheet_issues_count(self):
        return len(self._employee_timesheet_issues)

    def get_employee_timesheet_worklog_data(self):
        return self._employee_timesheet_worklog_data
