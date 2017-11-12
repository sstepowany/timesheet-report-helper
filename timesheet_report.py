import xlsxwriter
import os

from timesheets.timesheet import Timesheet
from dataProviders.timesheet_data_provider import TimesheetData
from utils.configuration_provider import Configuration
from utils.names_generator import NamesGenerator
from teamStructure.team_builder import ProjectTeamsBuilder
from utils.logger import TimesheetLogger
from utils.data_converter import DataConverter


class TimesheetReport(object):

    def __init__(self):
        self._teams_container = None
        self._timesheet_report_configuration = Configuration().get_timesheet_configuration()
        self._jira_configuration = Configuration().get_jira_configuration()
        self._timesheet_date = self._timesheet_report_configuration['timesheet_date']
        self._timesheet_data = TimesheetData(self._jira_configuration, self._timesheet_date)
        self._national_holidays_list = DataConverter().split_configuration_special_days_to_list(
            self._timesheet_date['national_holidays'])
        self._names_generator = NamesGenerator(self._timesheet_date)
        self._timesheet_reports_folder = 'reports'

    def create_timesheet(self, project_configuration, reports_set_folder, timesheet_name_prefix):
        TimesheetLogger().log_on_console("Started creating timesheet: {}.".format(timesheet_name_prefix))
        report_name = self._names_generator.generate_report_name(timesheet_name_prefix)
        if not os.path.exists(os.path.join(self._timesheet_reports_folder, reports_set_folder)):
            os.makedirs(os.path.join(self._timesheet_reports_folder, reports_set_folder))
        workbook = xlsxwriter.Workbook(os.path.join(self._timesheet_reports_folder, reports_set_folder, report_name))
        timesheet = Timesheet(workbook)
        self._teams_container = ProjectTeamsBuilder().build_project_team(
            project_configuration['teams'], self._timesheet_data)
        teams_list = self._teams_container.get_teams_list()
        timesheet.add_summary_worksheet(teams_list, self._timesheet_date, self._national_holidays_list,
                                        project_configuration['po_number'], project_configuration['man_days_costs'])
        timesheet.add_statistics_worksheet(teams_list, self._timesheet_date)
        timesheet.add_worksheets_for_employees(teams_list, self._timesheet_date)
        timesheet.save_timesheet()
        TimesheetLogger().log_on_console("Timesheet: {} created.".format(timesheet_name_prefix))

    def create_timesheets(self):
        TimesheetLogger().log_on_console("Started creating timesheets set.")
        reports_set_folder = self._names_generator.generate_timesheets_set_folder_name()
        for project in self._timesheet_report_configuration['timesheets']:
            self.create_timesheet(self._timesheet_report_configuration['timesheets'][project], reports_set_folder,
                                  project)
        TimesheetLogger().log_on_console("Timesheets set created.")
