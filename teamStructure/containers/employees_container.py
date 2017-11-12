from teamStructure.containers.base_container import BaseContainer
from teamStructure.employee import Employee


class EmployeesContainer(BaseContainer):

    def __init__(self):
        super(EmployeesContainer, self).__init__()

    def add_employee(self, employee_jira_name, employee_configuration_data, timesheet_data_object):
        self._list.append(Employee(employee_jira_name, employee_configuration_data, timesheet_data_object))

    def get_employees_list(self):
        return self.get_list()

    def get_employees_list_count(self):
        return self.get_list_items_count()
