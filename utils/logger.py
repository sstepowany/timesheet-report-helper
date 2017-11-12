import logging

from utils.date import Date


class TimesheetLogger(object):

    def __init__(self):
        pass

    _date = Date()
    _log_template = "{} - {}"
    _time_log_format = "%H:%M:%S"
    _logger = logging.getLogger('CONSOLE')
    _logger.setLevel(logging.DEBUG)
    _handler = logging.StreamHandler()
    _handler.setLevel(logging.DEBUG)
    _logger.addHandler(_handler)

    def log_on_console(self, log_text):
        log_time_raw = self._date.get_current_date_time()
        log_time = log_time_raw.strftime(self._time_log_format)
        log_text = self._log_template.format(log_time, log_text)
        self._logger.info(log_text)
