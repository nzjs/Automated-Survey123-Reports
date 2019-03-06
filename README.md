# Automated Survey123 Reports and Emailing

## What is it?
A tool to automate Survey123 Report creation & emailing survey recipients with a copy of the report.

This script takes an input Survey ID (from Survey123), a report template (uploaded to the Survey Reports (Beta) tab), and will do the following:

- Generate .docx reports for any new Survey submissions from the last 24 hours,
- Save these reports in bulk to AGOL, download them one by one to a location, then remove them from AGOL when finished,
- Read and extract an email address from each .docx file, 
- Send the relevant .docx file as an attachment to the relevant recipient, 
- Remove the .docx file once the email has sent,
- Logs the daily results to a txt file in the output folder.

Call this script with     python "..\S123ReportAndEmailSubmissions.py"     
- eg. in a Unix Cron job or Windows Task Scheduler that runs once a day - since we are always looking for submissions from the last 24 hours

Note: this script generates the KeyError: 'results' but still works due to the try/except/finally block...
Related to this ESRI bug:
*BUG-000119057 : The Python API 1.5.2 generate_report() method of the arcgis.apps.survey123 module, generates the following error: { KeyError: 'results' }*

API docs: https://esri.github.io/arcgis-python-api/apidoc/html/arcgis.apps.survey123.html

## Requirements
- Customise the .py file with the below variables as desired
- Use with Python v3 and install the required Python libraries
- Set the script to run daily with a Unix Cron job or Windows Task Scheduler

## Customisation
```python
# --- AGOL information... ---
org = 'https://YOUR-ORGANISATION.maps.arcgis.com'
username = 'ARCGIS ONLINE USERNAME'
password = 'ARCGIS ONLINE PASSWORD'


# --- Survey123 variables... ---
surveyID = 'ID OF SURVEY123 FORM' # ID of desired Survey123 form - a unique ID like 4c1b359c4e294c54a02b22b42413f1
output_folder = r'C:\GISWORK\_tmp\Reports' # Output folder WITHOUT trailing slash. This is also where the log file is stored.

# WHERE_FILTER: Use '1=1' to return all records, or something like  {{"where":"<col>='<value>'"}  - supports SQL syntax
# Docs for date queries: https://www.esri.com/arcgis-blog/products/api-rest/data-management/querying-feature-services-date-time-queries/
# In our case below, we filter by records created in the last 1 day (24 hrs). This works for us as the script is run on a daily schedule.
where_filter = '{"where":"CreationDate >= CURRENT_TIMESTAMP - INTERVAL \'1\' DAY"}'

utc_offset = '+13:00' # UTC Offset for location (+13 is NZST)
report_title = 'Daily_Export' # Title that will show in Survey123 Reports recent task list
report_template = 1 # ID of the print template in Survey123 that you want to use (0 = ESRI's sample, 1 = first custom report, 2 = second custom report, etc)


# --- Email SMTP settings... ---
email_user = 'EMAIL ADDRESS' # Eg. user@gmail.com. Requires a valid SMTP-enabled email account (Eg. a Gmail acct with the SMTP settings below)
email_password = 'EMAIL ACCOUNT PASSWORD' # Password for the email account
smtp_server = 'smtp.gmail.com'
smtp_port = 587
```
