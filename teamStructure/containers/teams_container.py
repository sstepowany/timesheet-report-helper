from teamStructure.containers.base_container import BaseContainer
from teamStructure.team import Team


class TeamsContainer(BaseContainer):

    def __init__(self):
        super(TeamsContainer, self).__init__()

    def add_team(self, team_name, employees_container):
        self._list.append(Team(team_name, employees_container))

    def get_teams_list(self):
        return self.get_list()
