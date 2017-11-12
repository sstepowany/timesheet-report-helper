from timesheets.worksheets.worksheetTemplates.statistics_worksheet_template import StatisticsWorksheetTemplate


class StatisticsWorksheet(StatisticsWorksheetTemplate):

    def __init__(self, formats_dict):
        self._formats_dict = formats_dict
        self._type_data_column_width = 15
        self._priority_data_column_width = 10
        self._component_data_column_width = 15
        self._team_member_data_column_width = 20
        self._priority_to_format_map = {
            'Blocker': 'blocker_priority_format',
            'Critical': 'critical_priority_format',
            'Major': 'major_priority_format',
            'Minor': 'minor_priority_format',
            'Trivial': 'trivial_priority_format'
        }
        super(StatisticsWorksheet, self).__init__()

    def prepare_statistics_worksheet(self, statistics_worksheet, teams_list, timesheet_date):
        timesheet_month = timesheet_date['month']
        timesheet_year = self._date.get_timesheet_year(timesheet_date['year'])
        teams_issues_count = 0
        for team in teams_list:
            teams_issues_count += team.get_team_issues_count()
        self.fill_statistics_worksheet_with_template(statistics_worksheet, self._formats_dict, timesheet_year,
                                                     timesheet_month, teams_issues_count)

        issues_count = 0
        for team in teams_list:
            for employee in team.get_employees_list():
                employee_worklog_data = employee.get_employee_timesheet_worklog_data()
                employee_issues = employee.get_employee_timesheet_issues()
                for issue in employee_issues:
                    issue_summary = issue['summary']
                    summary_field_length = len(issue_summary)
                    if summary_field_length > self._summary_field_longest_length:
                        self._summary_field_longest_length = summary_field_length

                    issue_key = issue['key']
                    issue_type = issue['issue_type']
                    issue_priority = issue['issue_priority']
                    issue_components = issue['issue_components']
                    employee_name = employee.get_employee_name()
                    issue_key_field_length = len(issue_key)
                    if issue_key_field_length > self._issue_key_field_longest_length:
                        self._issue_key_field_longest_length = issue_key_field_length

                    issue_row = issues_count + self._first_data_row_index
                    statistics_worksheet.write(issue_row, self._number_data_column_index, issues_count + 1)
                    statistics_worksheet.write(issue_row, self._summary_data_column_index, issue_summary)
                    statistics_worksheet.write(issue_row, self._task_id_data_column_index, issue_key)
                    statistics_worksheet.write(issue_row, self._type_data_column_index, issue_type)
                    statistics_worksheet.write(issue_row, self._priority_data_column_index, issue_priority,
                                               self._formats_dict[self._priority_to_format_map[issue_priority]])
                    statistics_worksheet.write(issue_row, self._component_data_column_index, issue_components,
                                               self._formats_dict['specific_background_color_format'])
                    statistics_worksheet.write(issue_row, self._team_member_data_column_index, employee_name,
                                               self._formats_dict['specific_background_color_format'])
                    issues_count += 1

                    statistics_worksheet.set_column(self._type_data_column_index, self._type_data_column_index,
                                                    self._type_data_column_width)
                    statistics_worksheet.set_column(self._priority_data_column_index, self._priority_data_column_index,
                                                    self._priority_data_column_width)
                    statistics_worksheet.set_column(self._component_data_column_index,
                                                    self._component_data_column_index,
                                                    self._component_data_column_width)
                    statistics_worksheet.set_column(self._team_member_data_column_index,
                                                    self._team_member_data_column_index,
                                                    self._team_member_data_column_width)

                    self.fill_user_worklog_data(statistics_worksheet, self._formats_dict, employee_worklog_data,
                                                self._first_month_day_column_index, issue_key, timesheet_year,
                                                timesheet_month, issue_row)

                self.fill_summary_section(statistics_worksheet, self._summary_row_index,
                                          self._summary_data_column_index, self._task_id_row_index,
                                          self._task_id_data_column_index)