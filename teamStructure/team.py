class Team(object):

    def __init__(self, team_name, employees_container):
        self._team_name = team_name
        self._employees_container = employees_container

    def get_employees_list(self):
        return self._employees_container.get_employees_list()

    def get_employees_list_count(self):
        return self._employees_container.get_employees_list_count()

    def get_team_name(self):
        return self._team_name

    def get_team_issues_count(self):
        team_issues_count = 0
        for employee in self._employees_container.get_employees_list():
            team_issues_count += len(employee.get_employee_timesheet_issues())
        return team_issues_count
