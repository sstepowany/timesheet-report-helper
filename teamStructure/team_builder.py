from teamStructure.containers.employees_container import EmployeesContainer
from teamStructure.containers.teams_container import TeamsContainer


class ProjectTeamsBuilder(object):

    def __init__(self):
        self._teams_container = TeamsContainer()

    def clean_teams_container(self):
        self._teams_container.clean_list()

    def build_project_team(self, timesheet_project_configuration, timesheet_data_object):
        for team in timesheet_project_configuration:
            employees_container = EmployeesContainer()
            for employee in timesheet_project_configuration[team]:
                employees_container.add_employee(
                    employee, timesheet_project_configuration[team][employee], timesheet_data_object)
            self._teams_container.add_team(team, employees_container)
        return self._teams_container
