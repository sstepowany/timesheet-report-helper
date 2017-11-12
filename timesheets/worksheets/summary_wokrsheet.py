from timesheets.worksheets.worksheetTemplates.summary_worksheet_template import SummaryWorksheetTemplate


class SummaryWorksheet(SummaryWorksheetTemplate):

    def __init__(self, formats_dict):
        self._formats_dict = formats_dict
        self._name_column_width = 15
        self._level_column_width = 15
        super(SummaryWorksheet, self).__init__()

    def prepare_summary_worksheet(self, summary_worksheet, teams_list, timesheet_date, national_holidays_days_list,
                                  po_number, man_days_costs):
        timesheet_month = timesheet_date['month']
        timesheet_year = self._date.get_timesheet_year(timesheet_date['year'])
        days_in_month = self._date.get_month_days_count(timesheet_year, timesheet_month)
        teams_and_employees_count = len(teams_list)
        for team in teams_list:
            teams_and_employees_count += team.get_employees_list_count()
        self.fill_summary_worksheet_with_template(summary_worksheet, self._formats_dict, timesheet_year,
                                                  timesheet_month, teams_and_employees_count, po_number, man_days_costs)

        row_index = 0
        for team in teams_list:
            team_name = team.get_team_name()
            team_row_index = self._first_data_row_index + row_index
            summary_worksheet.merge_range(team_row_index, self._name_data_column_index,
                                          team_row_index, self._first_month_day_column_index + days_in_month + 1,
                                          team_name, self._formats_dict['team_line_separator_format'])
            row_index += 1
            for employee in team.get_employees_list():
                employee_name = employee.get_employee_name()
                employee_level = employee.get_employee_configuration_level()
                employee_issues_count = employee.get_employee_timesheet_issues_count()
                employee_row_index = self._first_data_row_index + row_index
                employee_holidays_and_sick_leaves_list = \
                    employee.get_employee_configuration_holidays_and_sick_leaves_days_list()
                employee_trainings_and_other_projects_list = \
                    employee.get_employee_configuration_trainings_and_other_projects_days_list()

                employee_name_lenght = len(employee_name)
                if employee_name_lenght > self._name_column_width:
                    self._name_column_width = employee_name_lenght

                summary_worksheet.write(employee_row_index, self._name_data_column_index, employee_name)
                summary_worksheet.write(employee_row_index, self._level_data_column_index, employee_level,
                                        self._formats_dict['employee_level_text_format'])

                for day in range(1, days_in_month + 1):
                    is_day_of_weekend = self._date.is_day_of_weekend(timesheet_year, timesheet_month, day)
                    if is_day_of_weekend:
                        cell_format = self._formats_dict['weekend_day_cell_format']
                    elif len(national_holidays_days_list) > 0 and str(day) in national_holidays_days_list:
                        cell_format = self._formats_dict['national_holidays_color_format']
                    elif len(employee_holidays_and_sick_leaves_list) > 0 \
                            and str(day) in employee_holidays_and_sick_leaves_list:
                        cell_format = self._formats_dict['holidays_and_sick_leaves_color_format']
                    elif len(employee_trainings_and_other_projects_list) > 0 \
                            and str(day) in employee_trainings_and_other_projects_list:
                        cell_format = self._formats_dict['trainings_and_others_color_format']
                    else:
                        cell_format = None

                    employee_total_hours_row_index = self._first_data_row_index + employee_issues_count
                    employee_total_man_days_row_index = employee_total_hours_row_index + 1

                    # employee_total_hours_per_day_formula
                    self.write_indirect_formula(summary_worksheet, employee_row_index, self._name_data_column_index,
                                                employee_total_hours_row_index, self._first_month_day_column_index +
                                                day, employee_row_index,
                                                self._employee_logged_time_per_day_first_data_column_index + day - 1,
                                                cell_format)

                    # employee_total_man_days_per_month_formula
                    self.write_indirect_formula(summary_worksheet, employee_row_index, self._name_data_column_index,
                                                employee_total_man_days_row_index, self._first_month_day_column_index +
                                                days_in_month + 1, employee_row_index,
                                                self._employee_logged_time_per_day_first_data_column_index +
                                                days_in_month + 1)

                # total_man_days_formula
                self.write_sum_formula(summary_worksheet, employee_row_index, self._first_month_day_column_index,
                                       employee_row_index, days_in_month + self._first_month_day_column_index - 1,
                                       employee_row_index, self._first_month_day_column_index + days_in_month)
                row_index += 1

        summary_worksheet.set_column(self._name_data_column_index, self._name_data_column_index,
                                     self._name_column_width)
        summary_worksheet.set_column(self._level_data_column_index, self._level_data_column_index, 
                                     self._level_column_width)
