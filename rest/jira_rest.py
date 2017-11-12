from requests import get

from utils.date import Date


class JiraRest(object):

    def __init__(self, jira_configuration):
        self._jira_configuration = jira_configuration

    @staticmethod
    def prepare_search_parameters_for_timesheet_data(employee_jira_name, year, month):
        from_date, to_date = Date().get_month_range(year, month)
        params = {
            'fields': 'priority,components,key,issuetype,priority,summary',
            'jql': 'worklogAuthor={0} AND worklogDate>={1} AND worklogDate<={2}'.
                format(employee_jira_name, from_date, to_date)
        }
        return params

    @staticmethod
    def perapare_params_for_user_rest(employee_jira_name):
        params = {'username': employee_jira_name}
        return params

    def search_for_data(self, params):
        url = "{}{}".format(self._jira_configuration['base_url'], self._jira_configuration['search_endpoint'])
        response = get(url, params=params, auth=(self._jira_configuration['authentication']['user'],
                                                 self._jira_configuration['authentication']['password']))
        return response

    def get_issue_worklog(self, issue_id):
        url = "{}{}".format(self._jira_configuration['base_url'],
                            self._jira_configuration['worklog_endpoint'].format(issue_id))
        response = get(url, auth=(self._jira_configuration['authentication']['user'],
                                  self._jira_configuration['authentication']['password']))
        return response

    def get_user_data(self, params):
        url = "{}{}".format(self._jira_configuration['base_url'], self._jira_configuration['user_endpoint'])
        response = get(url, params=params, auth=(self._jira_configuration['authentication']['user'],
                                                 self._jira_configuration['authentication']['password']))
        return response
