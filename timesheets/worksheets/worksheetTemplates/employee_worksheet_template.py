from timesheets.worksheets.worksheetTemplates.base_worksheet_template import BaseWorksheetTemplate


class EmployeeWorksheetTemplate(BaseWorksheetTemplate):

    def __init__(self):
        self._table_columns_names_list = ['No.', 'Summary', 'Task ID']
        super(EmployeeWorksheetTemplate, self).__init__(self._table_columns_names_list)

    def fill_employee_worksheet_with_template(self, employee_worksheet, formats_dict, timesheet_year, timesheet_month,
                                              issues_count):
        if issues_count != 0:
            rows_count = issues_count
        else:
            rows_count = 1
        self.fill_worksheet_with_template(employee_worksheet, formats_dict, timesheet_year, timesheet_month, rows_count,
                                          self._task_id_data_column_index)
