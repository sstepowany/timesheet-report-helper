from xlsxwriter.utility import xl_range
from xlsxwriter.utility import xl_rowcol_to_cell

from utils.date import Date
from timesheets.worksheets.worksheetObjects.worksheet_cell import WorksheetCell
from timesheets.worksheets.worksheetObjects.worksheet_cells_range import WorksheetCellsRange
from timesheets.worksheets.worksheetObjects.worksheet_columns_range import WorksheetColumnsRange


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

    @staticmethod
    def write_to_cell(worksheet, cell):
        cell_row_index, cell_column_index = cell.get_cell_coordinates()
        cell_content = cell.get_cell_content()
        cell_format = cell.get_cell_format()
        worksheet.write(cell_row_index, cell_column_index, cell_content, cell_format)

    @staticmethod
    def write_to_blank_cell(worksheet, cell):
        cell_row_index, cell_column_index = cell.get_cell_coordinates()
        cell_content = ''
        cell_format = cell.get_cell_format()
        worksheet.write_blank(cell_row_index, cell_column_index, cell_content, cell_format)

    @staticmethod
    def set_columns(worksheet, columns_range):
        first_column_index, last_column_index = columns_range.get_range_columns_coordinates()
        first_column = columns_range.get_first_column()
        columns_width = first_column.get_column_width()
        worksheet.set_column(first_column_index, last_column_index, columns_width)

    @staticmethod
    def merge_range(worksheet, cells_range):
        first_range_cell_row_index, first_range_cell_column_index, second_range_cell_row_index, \
            second_range_cell_column_index = cells_range.get_range_cells_coordinates()
        range_second_cell = cells_range.get_range_second_cell()
        range_content = range_second_cell.get_cell_content()
        range_format = range_second_cell.get_cell_format()
        worksheet.merge_range(first_range_cell_row_index, first_range_cell_column_index,
                              second_range_cell_row_index, second_range_cell_column_index,
                              range_content, range_format)

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
                        time_spent_cell = WorksheetCell(issue_row, worklog_day + first_month_day_column_index - 1,
                                                        worklog['time_spent'], cell_format)
                        self.write_to_cell(worksheet, time_spent_cell)
                    else:
                        blank_time_spent_cell = WorksheetCell(issue_row, day + first_month_day_column_index - 1, '',
                                                              cell_format)
                        self.write_to_blank_cell(worksheet, blank_time_spent_cell)

        total_time_per_task_formula_range = \
            WorksheetCellsRange(issue_row, first_month_day_column_index, issue_row,
                                timesheet_last_month_day + first_month_day_column_index - 1, None, None)
        total_time_per_task_formula_cell = WorksheetCell(issue_row, timesheet_last_month_day +
                                                         first_month_day_column_index, None, None)
        self.write_sum_formula(worksheet, total_time_per_task_formula_range, total_time_per_task_formula_cell, None)

    def fill_summary_section(self, worksheet, summary_row_index, summary_column_index, task_id_row_index,
                             task_id_column_index):
        summary_columns_range = WorksheetColumnsRange(summary_row_index, summary_column_index,
                                                      self._summary_field_longest_length)
        self.set_columns(worksheet, summary_columns_range)
        task_id_columns_range = \
            WorksheetColumnsRange(task_id_row_index, task_id_column_index, self._issue_key_field_longest_length +
                                  self._issue_key_field_length_buffer)
        self.set_columns(worksheet, task_id_columns_range)

    def write_sum_formula_in_merge(self, worksheet, formula_range, formula_cells_range, cells_format=None):
        first_cell_row, first_cell_column, second_cell_row, second_cell_column = \
            formula_range.get_range_cells_coordinates()
        formula_first_cell_row, formula_first_cell_column, formula_second_cell_row, formula_second_cell_column = \
            formula_cells_range.get_range_cells_coordinates()
        formula_range = xl_range(first_cell_row, first_cell_column, second_cell_row, second_cell_column)
        formula = self._sum_formula_template.format(formula_range)
        worksheet.merge_range(formula_first_cell_row, formula_first_cell_column, formula_second_cell_row,
                              formula_second_cell_column, formula, cells_format)

    def write_sum_formula(self, worksheet, formula_cells_range, formula_cell, cells_format=None):
        first_cell_row, first_cell_column, second_cell_row, second_cell_column = \
            formula_cells_range.get_range_cells_coordinates()
        formula_cell_row, formula_cell_column = formula_cell.get_cell_coordinates()
        formula_range = xl_range(first_cell_row, first_cell_column, second_cell_row, second_cell_column)
        formula = self._sum_formula_template.format(formula_range)
        worksheet.write_formula(formula_cell_row, formula_cell_column, formula, cells_format)

    def write_indirect_formula(self, worksheet, first_cell, second_cell, formula_cell, cells_format=None):
        first_cell_row, first_cell_column = first_cell.get_cell_coordinates()
        second_cell_row, second_cell_column = second_cell.get_cell_coordinates()
        formula_cell_row, formula_cell_column = formula_cell.get_cell_coordinates()
        formula_first_cell = xl_rowcol_to_cell(first_cell_row, first_cell_column)
        formula_second_cell = xl_rowcol_to_cell(second_cell_row, second_cell_column)
        formula = self._indirect_formula_template.format(formula_first_cell, formula_second_cell)
        worksheet.write_formula(formula_cell_row, formula_cell_column, formula, cells_format)

    def write_sumif_formula(self, worksheet, first_formula_range, condition_cell, second_formula_range,
                            formula_cell, cells_format=None):
        formula_first_first_cell_row, formula_first_first_cell_column, formula_first_second_cell_row, \
        formula_first_second_cell_column = first_formula_range.get_range_cells_coordinates()

        formula_second_first_cell_row, formula_second_first_cell_column, formula_second_second_cell_row, \
        formula_second_second_cell_column = second_formula_range.get_range_cells_coordinates()

        condition_cell_row, condition_cell_column = condition_cell.get_cell_coordinates()
        formula_cell_row, formula_cell_column = formula_cell.get_cell_coordinates()

        first_range = xl_range(formula_first_first_cell_row, formula_first_first_cell_column,
                               formula_first_second_cell_row, formula_first_second_cell_column)
        second_range = xl_range(formula_second_first_cell_row, formula_second_first_cell_column,
                                formula_second_second_cell_row, formula_second_second_cell_column)
        formula_condition_cell = xl_rowcol_to_cell(condition_cell_row, condition_cell_column)
        formula = self._sumif_formula_template.format(first_range, formula_condition_cell, second_range)
        worksheet.write_formula(formula_cell_row, formula_cell_column, formula, cells_format)

    def write_product_formula_in_merge(self, worksheet, formula_cell, second_formula_parameter,
                                       formula_cells_range, cells_format=None):
        formula_cell_row, formula_cell_column = formula_cell.get_cell_coordinates()
        converted_formula_cell = xl_rowcol_to_cell(formula_cell_row, formula_cell_column)

        formula_first_cell_row, formula_first_cell_column, formula_second_cell_row, formula_second_cell_column =\
            formula_cells_range.get_range_cells_coordinates()

        formula = self._product_formula_template.format(converted_formula_cell, second_formula_parameter)
        worksheet.merge_range(formula_first_cell_row, formula_first_cell_column, formula_second_cell_row,
                              formula_second_cell_column, formula, cells_format)

    def write_man_days_counting_formula(self, worksheet, formula_cells_range, formula_cell, cells_format=None):
        formula_first_cell_row, formula_first_cell_column, formula_second_cell_row, formula_second_cell_column = \
            formula_cells_range.get_range_cells_coordinates()

        converted_formula_cells_range = xl_range(formula_first_cell_row, formula_first_cell_column,
                                                 formula_second_cell_row, formula_second_cell_column)
        formula_cell_row, formula_cell_column = formula_cell.get_cell_coordinates()
        formula = self._man_days_counting_formula_template.format(converted_formula_cells_range)
        worksheet.write_formula(formula_cell_row, formula_cell_column, formula, cells_format)
