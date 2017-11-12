from timesheets.worksheets.worksheetTemplates.base_worksheet_template import BaseWorksheetTemplate


class StatisticsWorksheetTemplate(BaseWorksheetTemplate):

    def __init__(self):
        self._table_columns_names_list = ['No.', 'Summary', 'Task ID', 'Type', 'Priority', 'Component', 'Team Memeber']
        self._type_data_column_index = 3
        self._priority_data_column_index = 4
        self._component_data_column_index = 5
        self._team_member_data_column_index = 6
        super(StatisticsWorksheetTemplate, self).__init__(self._table_columns_names_list)

    def fill_statistics_worksheet_with_template(self, statistics_worksheet, formats_dict, timesheet_year,
                                                timesheet_month, issues_count):
        self.fill_worksheet_with_template(statistics_worksheet, formats_dict, timesheet_year, timesheet_month,
                                          issues_count, self._team_member_data_column_index)
