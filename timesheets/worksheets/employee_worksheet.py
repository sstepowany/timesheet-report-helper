from timesheets.worksheets.worksheetTemplates.employee_worksheet_template import EmployeeWorksheetTemplate
from timesheets.worksheets.worksheetObjects.worksheet_cell import WorksheetCell


class EmployeeWorksheet(EmployeeWorksheetTemplate):

    def __init__(self, formats_dict):
        self._formats_dict = formats_dict
        super(EmployeeWorksheet, self).__init__()

    def prepare_worksheet_for_employee(self, employee_worksheet, employee_timesheet_issues, employee_worklog_data,
                                       timesheet_date, man_days_hours):
        timesheet_month = timesheet_date['month']
        timesheet_year = self._date.get_timesheet_year(timesheet_date['year'])
        issues_count = len(employee_timesheet_issues)
        self.fill_employee_worksheet_with_template(employee_worksheet, self._formats_dict, timesheet_year,
                                                   timesheet_month, issues_count, man_days_hours)

        for issue in employee_timesheet_issues:
            issue_summary = issue['summary']
            summary_field_length = len(issue_summary)
            if summary_field_length > self._summary_field_longest_length:
                self._summary_field_longest_length = summary_field_length

            issue_key = issue['key']
            issue_key_field_length = len(issue_key)
            if issue_key_field_length > self._issue_key_field_longest_length:
                self._issue_key_field_longest_length = issue_key_field_length

            issues_index = employee_timesheet_issues.index(issue)
            issue_row = issues_index + self._first_data_row_index

            issue_index_cell = WorksheetCell(issue_row, self._number_data_column_index, issues_index + 1, None)
            self.write_to_cell(employee_worksheet, issue_index_cell)

            issue_summary_cell = WorksheetCell(issue_row, self._summary_data_column_index, issue_summary, None)
            self.write_to_cell(employee_worksheet, issue_summary_cell)

            issue_key_cell = WorksheetCell(issue_row, self._task_id_data_column_index, issue_key, None)
            self.write_to_cell(employee_worksheet, issue_key_cell)

            self.fill_user_worklog_data(employee_worksheet, self._formats_dict, employee_worklog_data,
                                        self._first_month_day_column_index, issue_key, timesheet_year,
                                        timesheet_month, issue_row)

        self.fill_summary_section(employee_worksheet, self._summary_row_index, self._summary_data_column_index,
                                  self._task_id_row_index, self._task_id_data_column_index)
