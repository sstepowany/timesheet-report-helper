from timesheets.worksheets.worksheetTemplates.base_worksheet_template import BaseWorksheetTemplate
from timesheets.worksheets.worksheetObjects.worksheet_cell import WorksheetCell
from timesheets.worksheets.worksheetObjects.worksheet_cells_range import WorksheetCellsRange


class SummaryWorksheetTemplate(BaseWorksheetTemplate):

    def __init__(self):
        self._table_columns_names_list = ['Name', 'Level']
        self._name_data_column_index = 0
        self._level_data_column_index = 1
        self._employee_logged_time_per_day_first_data_column_index = 2
        self._purchase_order_row_index = 0
        self._summary_table_initial_column_index = 0
        self._summary_table_row_index_buffer = 6
        self._legend_row_index_buffer = 6
        self._legend_column_index_buffer = 5
        self._legend_header_cells_merge_range = 6
        self._purchase_order_text = "PO Number:"
        self._man_days_summary_table_summary_column_text = "Summary"
        self._man_days_summary_table_total_man_days_column_text = "Total man-days"
        self._man_days_summary_table_cost_column_text = "Cost"
        self._man_days_summary_table_total_row_text = "Total:"
        self._man_days_summary_table_levels_list = ['Senior', 'Medium', 'Junior']
        self._legend_title_text = "Legend:"
        self._legend_items_to_formula_map = {'National holidays': 'national_holidays_color_format',
                                             'Holidays / Sick leaves': 'holidays_and_sick_leaves_color_format',
                                             'Trainings / other project etc': 'trainings_and_others_color_format'}
        super(SummaryWorksheetTemplate, self).__init__(self._table_columns_names_list)

    def fill_summary_worksheet_with_template(self, summary_worksheet, formats_dict, timesheet_year, timesheet_month,
                                             teams_and_employees_count, po_number, man_days_costs):
        self.fill_worksheet_with_template(summary_worksheet, formats_dict, timesheet_year, timesheet_month,
                                          teams_and_employees_count, self._level_data_column_index,
                                          include_man_days_data=False, include_total_per_task_section=False)
        timesheet_last_month_day = self._date.get_month_days_count(timesheet_year, timesheet_month)

        purchase_order_cell = \
            WorksheetCell(self._purchase_order_row_index, self._name_data_column_index,
                          self._purchase_order_text, formats_dict['centered_and_bold_format'])
        self.write_to_cell(summary_worksheet, purchase_order_cell)
        po_number_cell = WorksheetCell(self._purchase_order_row_index + 1, self._name_data_column_index,
                                       po_number, formats_dict['centered_and_bold_format'])
        self.write_to_cell(summary_worksheet, po_number_cell)

        summary_total_man_days_column_index = self._summary_table_initial_column_index + 1
        summary_table_base_row_position_index = teams_and_employees_count + self._summary_table_row_index_buffer
        last_team_member_row_index = teams_and_employees_count + self._first_data_row_index - 1
        total_per_day_row_index = last_team_member_row_index + 1
        summary_table_cost_column_index = summary_total_man_days_column_index + 1
        total_per_person_summary_column_index = self._first_month_day_column_index + timesheet_last_month_day
        total_man_days_summary_column_index = total_per_person_summary_column_index + 1
        total_man_days_cost_column_index = summary_total_man_days_column_index + 1

        summary_table_summary_cell = \
            WorksheetCell(summary_table_base_row_position_index, self._summary_table_initial_column_index,
                          self._man_days_summary_table_summary_column_text,
                          formats_dict['summary_man_days_table_headers_format'])
        self.write_to_cell(summary_worksheet, summary_table_summary_cell)

        summary_table_man_days_cell = \
            WorksheetCell(summary_table_base_row_position_index, summary_total_man_days_column_index,
                          self._man_days_summary_table_total_man_days_column_text,
                          formats_dict['summary_man_days_table_headers_format'])
        self.write_to_cell(summary_worksheet, summary_table_man_days_cell)

        summary_table_cost_cells_range = \
            WorksheetCellsRange(summary_table_base_row_position_index, summary_table_cost_column_index,
                                summary_table_base_row_position_index, summary_table_cost_column_index + 1,
                                self._man_days_summary_table_cost_column_text,
                                formats_dict['summary_man_days_table_headers_format'])
        self.merge_range(summary_worksheet, summary_table_cost_cells_range)

        for level in self._man_days_summary_table_levels_list:
            level_index = self._man_days_summary_table_levels_list.index(level)
            if level_index % 2 == 1:
                cell_format = formats_dict['specific_background_color_with_align_center_format']
                formula_cell_format = formats_dict['centered_text_with_specific_color_format']
            else:
                cell_format = formats_dict['centered_text_format']
                formula_cell_format = formats_dict['centered_text_format']

            level_row_index = summary_table_base_row_position_index + level_index + 1
            level_cell = WorksheetCell(level_row_index, self._summary_table_initial_column_index, level, cell_format)
            self.write_to_cell(summary_worksheet, level_cell)

            total_man_days_column_index = self._first_data_row_index + timesheet_last_month_day
            total_man_days_formula_first_range = \
                WorksheetCellsRange(self._first_data_row_index, self._level_data_column_index,
                                    last_team_member_row_index, self._level_data_column_index, None, None)
            total_man_days_formula_condition_cell = \
                WorksheetCell(level_row_index,  self._summary_table_initial_column_index, None, None)
            total_man_days_formula_second_range = \
                WorksheetCellsRange(self._first_data_row_index, total_man_days_column_index, last_team_member_row_index,
                                    total_man_days_column_index, None, None)
            total_man_days_formula_cell = \
                WorksheetCell(level_row_index, summary_total_man_days_column_index, None, None)

            self.write_sumif_formula(summary_worksheet, total_man_days_formula_first_range,
                                     total_man_days_formula_condition_cell, total_man_days_formula_second_range,
                                     total_man_days_formula_cell, formula_cell_format)

            man_days_per_level_cost_formula_cell = \
                WorksheetCell(level_row_index, summary_total_man_days_column_index, None, None)
            man_days_per_level_cost_formula_range = \
                WorksheetCellsRange(level_row_index, summary_table_cost_column_index, level_row_index,
                                    summary_table_cost_column_index + 1, None, None)
            self.write_product_formula_in_merge(summary_worksheet, man_days_per_level_cost_formula_cell,
                                                man_days_costs[level], man_days_per_level_cost_formula_range,
                                                cell_format)

        levels_count = len(self._man_days_summary_table_levels_list)
        summary_table_total_row_index = summary_table_base_row_position_index + levels_count + 1
        first_level_row_index = summary_table_base_row_position_index + 1
        summary_table_table_total_cell = \
            WorksheetCell(summary_table_total_row_index, self._summary_table_initial_column_index,
                          self._man_days_summary_table_total_row_text, formats_dict['summary_total_row_format'])
        self.write_to_cell(summary_worksheet, summary_table_table_total_cell)

        total_man_days_summary_table_formula_range = WorksheetCellsRange(first_level_row_index, summary_total_man_days_column_index,
                               summary_table_base_row_position_index + levels_count,
                               summary_total_man_days_column_index, None, None)
        total_man_days_summary_table_formula_cell = WorksheetCell(summary_table_total_row_index, summary_total_man_days_column_index, None, None)

        self.write_sum_formula(summary_worksheet, total_man_days_summary_table_formula_range,
                               total_man_days_summary_table_formula_cell,
                               formats_dict['centered_and_bold_text_format'])

        # total_man_days_cost_formula
        levels_cells_range = \
            WorksheetCellsRange(first_level_row_index, total_man_days_cost_column_index,
                                summary_table_base_row_position_index + levels_count, total_man_days_cost_column_index,
                                None, None)
        formula_cells_range = WorksheetCellsRange(summary_table_total_row_index, summary_table_cost_column_index,
                                                  summary_table_total_row_index, summary_table_cost_column_index + 1,
                                                  None, None)
        self.write_sum_formula_in_merge(summary_worksheet, levels_cells_range, formula_cells_range,
                                        formats_dict['centered_and_bold_text_format'])

        total_per_person_summary_column_formula_range = \
            WorksheetCellsRange(self._first_data_row_index, total_per_person_summary_column_index,
                               last_team_member_row_index, total_per_person_summary_column_index, None, None)
        total_per_person_summary_column_formula_cell = WorksheetCell(total_per_day_row_index,
                                                                     total_per_person_summary_column_index, None, None)
        self.write_sum_formula(summary_worksheet, total_per_person_summary_column_formula_range,
                               total_per_person_summary_column_formula_cell, formats_dict['table_header_format'])

        total_man_days_summary_formula_range = \
            WorksheetCellsRange(self._first_data_row_index, total_man_days_summary_column_index,
                                last_team_member_row_index, total_man_days_summary_column_index, None, None)
        total_man_days_summary_formula_cell = \
            WorksheetCell(total_per_day_row_index, total_man_days_summary_column_index, None, None)
        self.write_sum_formula(summary_worksheet, total_man_days_summary_formula_range,
                               total_man_days_summary_formula_cell, formats_dict['table_header_format'])

        legend_base_row_position_index = teams_and_employees_count + self._legend_row_index_buffer
        legend_base_column_index = self._summary_table_initial_column_index + self._legend_column_index_buffer

        legend_header_cells_range = \
            WorksheetCellsRange(legend_base_row_position_index, legend_base_column_index,
                                legend_base_row_position_index, legend_base_column_index +
                                self._legend_header_cells_merge_range, self._legend_title_text,
                                formats_dict['centered_and_bold_format'])
        self.merge_range(summary_worksheet, legend_header_cells_range)

        legend_item_index = 0
        for legend_item in self._legend_items_to_formula_map:
            legend_item_cells_range = \
                WorksheetCellsRange(legend_base_row_position_index + legend_item_index + 1,
                                    legend_base_column_index + 1, legend_base_row_position_index +
                                    legend_item_index + 1, legend_base_column_index +
                                    self._legend_header_cells_merge_range, legend_item, None)
            self.merge_range(summary_worksheet, legend_item_cells_range)

            legend_item_color_cell = WorksheetCell(legend_base_row_position_index + legend_item_index + 1,
                                                   legend_base_column_index, '',
                                                   formats_dict[self._legend_items_to_formula_map[legend_item]])
            self.write_to_blank_cell(summary_worksheet, legend_item_color_cell)
            legend_item_index += 1
