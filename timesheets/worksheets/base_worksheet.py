from xlsxwriter.utility import xl_range
from xlsxwriter.utility import xl_rowcol_to_cell

from utils.date import Date


class BaseWorksheet(object):

    def __init__(self):
        self._date = Date()
        self._summary_field_longest_length = 5
        self._issue_key_field_longest_length = 5
        self._issue_key_field_length_buffer = 2
        self._sum_formula_template = '=SUM({})'
        self._indirect_formula_template = "=INDIRECT(\"'\"&{}&\"'!{}\")"
        self._sumif_formula_template = '=SUMIF({},{},{})'
        self._product_formula_template = '=PRODUCT({},{})'
        self._man_days_counting_formula_template = "={}/7"

    def fill_user_worklog_data(self, worksheet, formats_dict, employee_worklog_data, first_month_day_column_index,
                               issue_key, timesheet_year, timesheet_month, issue_row):
        timesheet_last_month_day = self._date.get_month_days_count(timesheet_year, timesheet_month)
        for worklog in employee_worklog_data[issue_key]:
            for day in range(1, timesheet_last_month_day + 1):
                worklog_year = worklog['worklog_date']['year']
                worklog_month = worklog['worklog_date']['month']
                if int(worklog_year) == timesheet_year and int(worklog_month) == timesheet_month:
                    worklog_day = int(worklog['worklog_date']['day'])
                    is_day_of_weekend = self._date.is_day_of_weekend(timesheet_year, timesheet_month, day)
                    if is_day_of_weekend:
                        cell_format = formats_dict['weekend_day_cell_format']
                    else:
                        cell_format = None

                    if worklog_day == day:
                        worksheet.write(issue_row, worklog_day + first_month_day_column_index - 1,
                                        worklog['time_spent'], cell_format)
                    else:
                        worksheet.write_blank(issue_row, day + first_month_day_column_index - 1, '',
                                              cell_format)

        self.write_sum_formula(worksheet, issue_row, first_month_day_column_index, issue_row,
                               timesheet_last_month_day + first_month_day_column_index - 1, issue_row,
                               timesheet_last_month_day + first_month_day_column_index)

    def fill_summary_section(self, worksheet, summary_row_index, summary_column_index, task_id_row_index,
                             task_id_column_index):
        worksheet.set_column(summary_row_index, summary_column_index, self._summary_field_longest_length)
        worksheet.set_column(task_id_row_index, task_id_column_index, self._issue_key_field_longest_length +
                             self._issue_key_field_length_buffer)


    def write_sum_formula_in_merge(self, worksheet, first_cell_row, first_cell_column, second_cell_row,
                                   second_cell_column, formula_first_cell_row, formula_first_cell_column,
                                   formula_second_cell_row, formula_second_cell_column, format=None):
        formula_range = xl_range(first_cell_row, first_cell_column, second_cell_row, second_cell_column)
        formula = self._sum_formula_template.format(formula_range)
        worksheet.merge_range(formula_first_cell_row, formula_first_cell_column, formula_second_cell_row,
                              formula_second_cell_column, formula, format)

    def write_sum_formula(self, worksheet, first_cell_row, first_cell_column, second_cell_row, second_cell_column,
                          formula_cell_row, formula_cell_column, format=None):
        formula_range = xl_range(first_cell_row, first_cell_column, second_cell_row, second_cell_column)
        formula = self._sum_formula_template.format(formula_range)
        worksheet.write_formula(formula_cell_row, formula_cell_column, formula, format)

    def write_indirect_formula(self, worksheet, first_cell_row, first_cell_column, second_cell_row, second_cell_column,
                          formula_cell_row, formula_cell_column, format=None):
        first_cell = xl_rowcol_to_cell(first_cell_row, first_cell_column)
        second_cell = xl_rowcol_to_cell(second_cell_row, second_cell_column)
        formula = self._indirect_formula_template.format(first_cell, second_cell)
        worksheet.write_formula(formula_cell_row, formula_cell_column, formula, format)

    def write_sumif_formula(self, worksheet, first_range_first_cell_row, first_range_first_cell_column,
                            first_range_second_cell_row, first_range_second_cell_column, condition_cell_row,
                            condition_cell_column, second_range_first_cell_row, second_range_first_cell_column,
                            second_range_second_cell_row, second_range_second_cell_column, formula_cell_row,
                            formula_cell_column, format=None):
        first_range = xl_range(first_range_first_cell_row, first_range_first_cell_column,
                               first_range_second_cell_row, first_range_second_cell_column)
        second_range = xl_range(second_range_first_cell_row, second_range_first_cell_column,
                                second_range_second_cell_row, second_range_second_cell_column)
        condition_cell = xl_rowcol_to_cell(condition_cell_row, condition_cell_column)
        formula = self._sumif_formula_template.format(first_range, condition_cell, second_range)
        worksheet.write_formula(formula_cell_row, formula_cell_column, formula, format)

    def write_product_formula(self, worksheet, cell_row, cell_column, second_formula_parameter, formula_cell_row,
                              formula_cell_column, format=None):
        formula_cell = xl_rowcol_to_cell(cell_row, cell_column)
        formula = self._product_formula_template.format(formula_cell, second_formula_parameter)
        worksheet.write_formula(formula_cell_row, formula_cell_column, formula, format)

    def write_product_formula_in_merge(self, worksheet, cell_row, cell_column, second_formula_parameter,
                                       formula_first_cell_row, formula_first_cell_column, formula_second_cell_row,
                                       formula_secind_cell_column, format=None):
        formula_cell = xl_rowcol_to_cell(cell_row, cell_column)
        formula = self._product_formula_template.format(formula_cell, second_formula_parameter)
        worksheet.merge_range(formula_first_cell_row, formula_first_cell_column, formula_second_cell_row,
                                formula_secind_cell_column, formula, format)

    def write_man_days_counting_formula(self, worksheet, first_cell_row, first_cell_column, second_cell_row,
                                        second_cell_column, formula_cell_row, formula_cell_column, format=None):
        formula_cells_range = xl_range(first_cell_row, first_cell_column, second_cell_row, second_cell_column)
        formula = self._man_days_counting_formula_template.format(formula_cells_range)
        worksheet.write_formula(formula_cell_row, formula_cell_column, formula, format)
