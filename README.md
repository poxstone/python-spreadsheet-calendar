# Get calendar month events
Put calendar

## Create project and client secret
- Create Project and **client_secret.json** file and put in self directory [Google QuickStart](https://developers.google.com/sheets/api/quickstart/python),
- Enable APIs:
  - Calendar
  - Spreedsheet
  - Drive

## Install dependencies
    pip install --upgrade httplib2 -t lib/
    pip install --upgrade google-api-python-client -t lib/
    pip install --upgrade python-dateutil -t lib/

## Execute
    python get_calendar_month_events.py
    