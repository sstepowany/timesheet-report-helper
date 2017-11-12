class TimesheetFormats(object):

    def __init__(self, workbook):
        self._workbook = workbook
        self._specific_background_color = '#eaeaea'
        self._centered_and_bold_format = self._workbook.add_format({
            'bold': True,
            'align': 'center',
            'border': 1
        })
        self._table_header_format = self._workbook.add_format({
            'bold': True,
            'bg_color': self._specific_background_color,
            'border': 0,
            'bottom': 1,
        })
        self._table_row_summary_section_aligned_right_format = self._workbook.add_format({
            'bold': True,
            'bg_color': self._specific_background_color,
            'border': 0,
            'bottom': 1,
            'align': 'right',
        })
        self._weekend_day_cell_format = self._workbook.add_format({
            'border': 0,
            'bg_color': self._specific_background_color
        })
        self._blocker_priority_format = self._workbook.add_format({
            'bg_color': '#e06666'
        })
        self._critical_priority_format = self._workbook.add_format({
            'bg_color': '#f7caac'
        })
        self._major_priority_format = self._workbook.add_format({
            'bg_color': '#ffe598'
        })
        self._minor_priority_format = self._workbook.add_format({
            'bg_color': '#c5e0b3'
        })
        self._trivial_priority_format = self._workbook.add_format({
            'bg_color': '#dee0f6'
        })
        self._specific_background_color_format = self._workbook.add_format({
            'bg_color': self._specific_background_color
        })
        self._specific_background_color_with_align_right_format = self._workbook.add_format({
            'bg_color': self._specific_background_color,
            'align': 'right'
        })
        self._team_line_separator_format = self._workbook.add_format({
            'bold': True,
            'bg_color': '#5b9bd5',
            'align': 'center',
            'color': '#f9f9f9'
        })
        self._summary_man_days_table_headers_format = self._workbook.add_format({
            'bold': True,
            'bg_color': '#5b9bd5',
            'align': 'center',
            'color': '#f9f9f9'
        })
        self._summary_total_row_format = self._workbook.add_format({
            'bold': True,
            'align': 'right',
            'top': 1
        })
        self._centered_text_format = self._workbook.add_format({
            'align': 'center'
        })
        self._centered_text_with_specific_color_format = self._workbook.add_format({
            'bg_color': self._specific_background_color,
            'align': 'center'
        })
        self._aligned_right_text_format = self._workbook.add_format({
            'align': 'right'
        })
        self._centered_and_bold_text_format = self._workbook.add_format({
            'bold': True,
            'align': 'center',
            'top': 1
        })
        self._national_holidays_color_format = self._workbook.add_format({
            'bg_color': '#f7caac'
        })
        self._holidays_and_sick_leaves_color_format = self._workbook.add_format({
            'bg_color': '#ffe598'
        })
        self._trainings_and_others_color_format = self._workbook.add_format({
            'bg_color': '#c5e0b3'
        })
        self._employee_level_text_format = self._workbook.add_format({
            'color': '#808080',
            'italic': True
        })
        self._table_row_summary_section_format = self._table_header_format
        self._formats_dict = {
            'centered_and_bold_format': self._centered_and_bold_format,
            'table_header_format': self._table_header_format,
            'table_row_summary_section_format': self._table_row_summary_section_format,
            'table_row_summary_section_aligned_right_format': self._table_row_summary_section_aligned_right_format,
            'weekend_day_cell_format': self._weekend_day_cell_format,
            'specific_background_color_format': self._specific_background_color_format,
            'blocker_priority_format': self._blocker_priority_format,
            'critical_priority_format': self._critical_priority_format,
            'major_priority_format': self._major_priority_format,
            'minor_priority_format': self._minor_priority_format,
            'trivial_priority_format': self._trivial_priority_format,
            'national_holidays_color_format': self._national_holidays_color_format,
            'team_line_separator_format': self._team_line_separator_format,
            'holidays_and_sick_leaves_color_format' : self._holidays_and_sick_leaves_color_format,
            'trainings_and_others_color_format': self._trainings_and_others_color_format,
            'summary_man_days_table_headers_format': self._summary_man_days_table_headers_format,
            'summary_total_row_format': self._summary_total_row_format,
            'centered_text_format': self._centered_text_format,
            'centered_and_bold_text_format': self._centered_and_bold_text_format,
            'specific_background_color_with_align_right_format': self._specific_background_color_with_align_right_format,
            'aligned_right_text_format': self._aligned_right_text_format,
            'centered_text_with_specific_color_format': self._centered_text_with_specific_color_format,
            'employee_level_text_format': self._employee_level_text_format
        }
