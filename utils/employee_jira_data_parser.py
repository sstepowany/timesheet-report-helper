from utils.date import Date
from utils.time import Time


class EmployeeJiraDataParser(object):

    def __init__(self):
        self._date = Date()
        self._time = Time()

    @staticmethod
    def _get_task_components_as_string(components_list):
        components_string = ""
        for component in components_list:
            if components_string != "":
                components_string = "{};{}".format(components_string, component['name'])
            else:
                components_string = component['name']
        return components_string

    def get_jira_issues(self, employee_timesheet_data):
        jira_issues = list()
        for issue in employee_timesheet_data['issues']:
            issue_dict = dict()
            issue_dict['summary'] = issue['fields']['summary']
            issue_dict['key'] = issue['key']
            issue_dict['issue_id'] = issue['id']
            issue_dict['issue_type'] = issue['fields']['issuetype']['name']
            issue_dict['issue_priority'] = issue['fields']['priority']['name']
            issue_dict['issue_components'] = self._get_task_components_as_string(issue['fields']['components'])
            jira_issues.append(issue_dict)
        return jira_issues

    def parse_employe_worklog_data(self, employee_worklog_data, employee_jira_name, issue_worklog_data):
        worklogs = issue_worklog_data['worklogs']
        for worklog in worklogs:
            if worklog['author']['name'] == employee_jira_name:
                worklog_data = dict()
                worklog_data['worklog_date'] = dict()
                worklog_year, worklog_month, worklog_day = self._date.parse_worklog_date_time(worklog['started'])
                worklog_data['worklog_date']['year'] = worklog_year
                worklog_data['worklog_date']['month'] = worklog_month
                worklog_data['worklog_date']['day'] = worklog_day
                time_spend_hours = float(self._time.parse_seconds_to_hours(worklog['timeSpentSeconds']))
                worklog_data['time_spent'] = time_spend_hours
                employee_worklog_data.append(worklog_data)
        return employee_worklog_data

    def parse_employe_worklog_data_2(self, employee_worklog_data, employee_jira_name, issue_worklog_data):
        worklogs = issue_worklog_data['worklogs']
        time_spent_per_date = dict()
        for worklog in worklogs:
            if worklog['author']['name'] == employee_jira_name:
                worklog_data = dict()
                worklog_data['worklog_date'] = dict()
                worklog_year, worklog_month, worklog_day = self._date.parse_worklog_date_time(worklog['started'])
                worklog_data['worklog_date']['year'] = worklog_year
                worklog_data['worklog_date']['month'] = worklog_month
                worklog_data['worklog_date']['day'] = worklog_day
                time_spend_hours = float(self._time.parse_seconds_to_hours(worklog['timeSpentSeconds']))
                time_spent_per_date_key = "{0}-{1}-{2}".format(worklog_year, worklog_month, worklog_day)
                if time_spent_per_date_key in time_spent_per_date.keys():
                    time_spent_per_date[time_spent_per_date_key] += time_spend_hours
                else:
                    time_spent_per_date[time_spent_per_date_key] = time_spend_hours
                worklog_data['time_spent'] = time_spent_per_date[time_spent_per_date_key]
                employee_worklog_data.append(worklog_data)
        return employee_worklog_data
