from timesheets.worksheets.base_worksheet import BaseWorksheet
from timesheets.worksheets.worksheetObjects.worksheet_cell import WorksheetCell
from timesheets.worksheets.worksheetObjects.worksheet_columns_range import WorksheetColumnsRange
from timesheets.worksheets.worksheetObjects.worksheet_cells_range import WorksheetCellsRange


class BaseWorksheetTemplate(BaseWorksheet):

    def __init__(self, table_columns_names_list):
        self._day_row_index = 2
        self._number_data_column_index = 0
        self._summary_data_column_index = 1
        self._task_id_data_column_index = 2
        self._year_row_index = 0
        self._summary_row_index = 1
        self._task_id_row_index = 2
        self._first_data_row_index = 3
        self._month_name_row_index = 1
        self._day_column_width = 4
        self._summary_column_width = 15
        self._table_columns_names_list = table_columns_names_list
        self._first_month_day_column_index = len(self._table_columns_names_list)
        self._summary_column_text = "Total per task"
        self._summary_per_person_column_text = "Total per person"
        self._summary_man_days_column_text = "Total man-days"
        self._summary_row_total_per_day_text = "Total per day:"
        self._summary_row_man_days_text = "Man-days:"
        super(BaseWorksheetTemplate, self).__init__()

    def fill_worksheet_with_template(self, worksheet, formats_dict, timesheet_year, timesheet_month, rows_count,
                                     last_table_basic_column_header_index, include_man_days_data=True,
                                     include_total_per_task_section=True):
        timesheet_last_month_day = self._date.get_month_days_count(timesheet_year, timesheet_month)
        self.fill_worksheet_with_basic_month_data(worksheet, formats_dict, self._table_columns_names_list,
                                                  timesheet_last_month_day, include_total_per_task_section)
        time_summary_row_index = rows_count + self._first_data_row_index
        if include_man_days_data:
            man_days_row_index = time_summary_row_index + 1
            self.fill_man_days_summary_section(worksheet, formats_dict, man_days_row_index, timesheet_last_month_day,
                                               last_table_basic_column_header_index)
        else:
            man_days_row_index = None
        self.fill_worksheet_with_month_days_specific_data(worksheet, formats_dict, timesheet_last_month_day, rows_count,
                                                          time_summary_row_index, man_days_row_index)
        self.fill_timesheet_date_data(worksheet, formats_dict, timesheet_month, timesheet_year,
                                      timesheet_last_month_day)
        self.fill_totals_counting_columns(worksheet, formats_dict, timesheet_last_month_day,
                                          include_total_per_task_section)

        self.fill_total_worklog_summary_section(worksheet, formats_dict, time_summary_row_index,
                                                timesheet_last_month_day, last_table_basic_column_header_index,
                                                include_total_per_task_section)

    def fill_worksheet_with_basic_month_data(self, worksheet, formats_dict, table_columns_names_list,
                                               timesheet_last_month_day, include_total_per_task_section):
        for column_name in table_columns_names_list:
            column_header_cell = WorksheetCell(self._day_row_index, table_columns_names_list.index(column_name),
                                               column_name, formats_dict['table_header_format'])
            self.write_to_cell(worksheet, column_header_cell)

        day_columns_range = WorksheetColumnsRange(self._first_month_day_column_index, timesheet_last_month_day +
                                                  self._first_month_day_column_index, self._day_column_width)
        self.set_columns(worksheet, day_columns_range)

        summary_columns_range = WorksheetColumnsRange(timesheet_last_month_day + self._first_month_day_column_index,
                             timesheet_last_month_day + self._first_month_day_column_index,
                             self._summary_column_width)
        self.set_columns(worksheet, summary_columns_range)

        if include_total_per_task_section is False:
            summary_total_per_tasks_columns_range = \
                WorksheetColumnsRange(timesheet_last_month_day + self._first_month_day_column_index + 1,
                                      timesheet_last_month_day + self._first_month_day_column_index + 1,
                                      self._summary_column_width)
            self.set_columns(worksheet, summary_total_per_tasks_columns_range)

    def fill_worksheet_with_month_days_specific_data(self, worksheet, formats_dict, timesheet_last_month_day,
                                                     issues_count, time_summary_row_index, man_days_row_index):
        for day in range(0, timesheet_last_month_day):
            month_day_column_index = day + self._first_month_day_column_index
            month_day_cell = WorksheetCell(self._day_row_index, month_day_column_index, day + 1,
                                           formats_dict['table_header_format'])
            self.write_to_cell(worksheet, month_day_cell)

            total_worklog_time_spent_for_day_formula_range = \
                WorksheetCellsRange(self._first_data_row_index, month_day_column_index,
                                   self._first_data_row_index + issues_count - 1, month_day_column_index, None, None)
            total_worklog_time_spent_for_day_formula_cell = WorksheetCell(time_summary_row_index,
                                                                          month_day_column_index, None, None)
            self.write_sum_formula(worksheet, total_worklog_time_spent_for_day_formula_range,
                                   total_worklog_time_spent_for_day_formula_cell, formats_dict['table_header_format'])
            if man_days_row_index is not None:
                man_days_counting_formula_range = \
                    WorksheetCellsRange(time_summary_row_index, month_day_column_index, time_summary_row_index,
                                        month_day_column_index, None, None)
                man_days_counting_formula_cell = \
                    WorksheetCell(man_days_row_index, month_day_column_index, None, None)
                self.write_man_days_counting_formula(worksheet, man_days_counting_formula_range,
                                                     man_days_counting_formula_cell,
                                                     formats_dict['table_row_summary_section_format'])

    def fill_timesheet_date_data(self, worksheet, formats_dict, timesheet_month, timesheet_year,
                                 timesheet_last_month_day):
        timesheet_month_name = self._date.get_month_name(timesheet_month)

        month_cells_range = \
            WorksheetCellsRange(self._month_name_row_index, self._first_month_day_column_index,
                                self._month_name_row_index, timesheet_last_month_day +
                                self._first_month_day_column_index - 1, timesheet_month_name,
                                formats_dict['centered_and_bold_format'])
        self.merge_range(worksheet, month_cells_range)
        year_cells_range = \
            WorksheetCellsRange(self._year_row_index, self._first_month_day_column_index,
                              self._year_row_index, timesheet_last_month_day +
                              self._first_month_day_column_index - 1, timesheet_year,
                              formats_dict['centered_and_bold_format'])
        self.merge_range(worksheet, year_cells_range)

    def fill_totals_counting_columns(self, worksheet, formats_dict, timesheet_last_month_day,
                                     include_total_per_task_section=True):
        if include_total_per_task_section:
            summary_columns_range = \
                WorksheetCellsRange(self._year_row_index, timesheet_last_month_day +
                                    self._first_month_day_column_index,
                                    self._day_row_index, timesheet_last_month_day +
                                    self._first_month_day_column_index,
                                    self._summary_column_text, formats_dict['centered_and_bold_format'])
            self.merge_range(worksheet, summary_columns_range)
        else:
            summary_per_person_columns_range = \
                WorksheetCellsRange(self._year_row_index, timesheet_last_month_day +
                                    self._first_month_day_column_index,
                                    self._day_row_index, timesheet_last_month_day +
                                    self._first_month_day_column_index,
                                    self._summary_per_person_column_text, formats_dict['centered_and_bold_format'])
            self.merge_range(worksheet, summary_per_person_columns_range)
            summary_man_days_columns_range = \
                WorksheetCellsRange(self._year_row_index, timesheet_last_month_day +
                                    self._first_month_day_column_index + 1,
                                    self._day_row_index, timesheet_last_month_day +
                                    self._first_month_day_column_index + 1,
                                    self._summary_man_days_column_text, formats_dict['centered_and_bold_format'])
            self.merge_range(worksheet, summary_man_days_columns_range)

    def fill_total_worklog_summary_section(self, worksheet, formats_dict, time_summary_row_index,
                                           timesheet_last_month_day, last_table_basic_column_header_index,
                                           include_total_per_task_section):
        summary_total_per_day_cells_range = \
            WorksheetCellsRange(time_summary_row_index, self._number_data_column_index,
                                time_summary_row_index, last_table_basic_column_header_index,
                                self._summary_row_total_per_day_text,
                                formats_dict['table_row_summary_section_aligned_right_format'])
        self.merge_range(worksheet, summary_total_per_day_cells_range)

        if include_total_per_task_section:
            total_worklog_per_month_formula_range = \
                WorksheetCellsRange(time_summary_row_index, self._first_month_day_column_index,
                                    time_summary_row_index, timesheet_last_month_day +
                                    self._first_month_day_column_index - 1, None, None)
            total_worklog_per_month_formula_cell = WorksheetCell(time_summary_row_index,
                                   timesheet_last_month_day + self._first_month_day_column_index, None, None)
            self.write_sum_formula(worksheet, total_worklog_per_month_formula_range,
                                   total_worklog_per_month_formula_cell,
                                   formats_dict['table_row_summary_section_aligned_right_format'])

    def fill_man_days_summary_section(self, worksheet, formats_dict, man_days_row_index, timesheet_last_month_day,
                                      last_table_basic_column_header_index):
        total_man_days_per_month_formula_range = \
            WorksheetCellsRange(man_days_row_index, self._first_month_day_column_index, man_days_row_index,
                               timesheet_last_month_day + self._first_month_day_column_index - 1, None, None)
        total_man_days_per_month_formula_cell = WorksheetCell(man_days_row_index,
                               timesheet_last_month_day + self._first_month_day_column_index, None, None)
        self.write_sum_formula(worksheet, total_man_days_per_month_formula_range,
                               total_man_days_per_month_formula_cell, formats_dict['table_header_format'])

        summary_man_days_cells_range = \
            WorksheetCellsRange(man_days_row_index, self._number_data_column_index,
                              man_days_row_index, last_table_basic_column_header_index, self._summary_row_man_days_text,
                              formats_dict['table_row_summary_section_aligned_right_format'])
        self.merge_range(worksheet, summary_man_days_cells_range)
