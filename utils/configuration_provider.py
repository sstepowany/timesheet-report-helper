import yaml
import os
import io


class Configuration(object):

    def __init__(self):
        self._configuration_folder = 'config'
        self._timesheet_configuration_file_path = os.path.join(self._configuration_folder, 'timesheet_config.yaml')
        self._jira_configuration_file_path = os.path.join(self._configuration_folder, 'jira_config.yaml')

    def get_timesheet_configuration(self):
        with io.open(self._timesheet_configuration_file_path, 'r', encoding='utf8') as config_file:
            timesheet_configuration = yaml.safe_load(config_file)
        return timesheet_configuration

    def get_jira_configuration(self):
        with io.open(self._jira_configuration_file_path, 'r', encoding='utf8') as config_file:
            jira_configuration = yaml.safe_load(config_file)
        return jira_configuration
