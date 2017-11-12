Operation manual

1. Install python in version 2.7.14, if not already installed.
1.1. Python installers are located in folder "wheels/pythonInstallers".
1.2. Choose correct installer depending on a windows version (32 or 64bit).
1.3. After installation make sure you have python executable added to the PATH variable.

2. Install dependencies for report helper.
2.1. Open command line tool.
2.2. Go to timesheet-report-helper location. (command ex. "cd d:/timesheet-report-heleper").
2.3. Run setup command: "python setup.py".
2.4. You can check if all dependencies from "requirements.txt" file were installed by writing command "pip list".

3. Configure the toll.
3.1. Timesheet helper has two configuration files.
3.2. First one is for communication with JIRA.
3.2.1. Copy "jira_config_template.yaml".
3.2.2. Name it as "jira_config.yaml".
3.2.3. Replace "user" with your jira id.
3.2.4. Replace "password" with your jira password.
3.3. Second one is for setting timesheet date and projects teams data.
3.3.1. Copy "timesheet_config_template.yaml".
3.3.2. Name it as "timesheet_config.yaml".
3.3.3. Replace all required values with you personalized timesheet report data.

4. Running the the timesheet-report-helper tool.
4.1. Open command line tool.
4.2. Go to timesheet-report-helper location. (command ex. "cd d:/timesheet-report-heleper").
4.3. Run command: "python generate_timesheet_report.py".
4.4. After timesheet generation process, created timesheets can be found in location: "reports/<specific_timesheet_date>".
