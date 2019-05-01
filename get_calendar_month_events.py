# -*- coding: utf-8 -*-
from __future__ import print_function
import sys
sys.path.insert(0, './lib')

import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
import httplib2
import dateutil.parser
import datetime
import calendar
import re

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

print(flags)

SCOPES = ['https://www.googleapis.com/auth/calendar.readonly',
          'https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive']
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Extract Calendar tasks'
PROJECTS = ['XSA', 'XSAU', 'XSC', 'XSCM', 'XSF', 'XSK', 'XSKIT', 'XSP', 'XSS',
            'XST', 'XSX', 'GCP', 'ALDEAMO', 'UNAL']
TASKS = ['DEV', 'DSN', 'ENT', 'IMP', 'INV', 'REU', 'SOP', 'DOC']
SHEET_NAME = 'DEV-OUTPUT-CALENDAR-2'
SHEET_PAGE = 'month_'
MONTH = 4
YEAR = None


def get_credentials():
    credential_path = os.path.join('.credentials.json')
    store = Storage(credential_path)
    credentials = store.get()
    
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)

    return credentials


def extractType(description, list_word):
    find = '-'

    for word in list_word:
        word_re = re.compile('.*' + word + '\W.*')
        if re.match(word_re, description.decode("utf-8")):
            find = word
            break

    return find


def prettyDate(date_string):
    try:
        date_string = date_string.decode("utf-8")
    except Exception as e:
        pritn('error')

    pretty_date = re.sub(r'^(\d+-){2}', '', date_string)
    pretty_date = re.sub(r'((:|-)\d+){3}$', '', pretty_date)
    return re.sub(r'T', '/', pretty_date)


def getPageIdByName(spread_sheet, page_name):
    pages = spread_sheet.get('sheets')
    page_id = ''

    for page in pages:
        current_page_name = page.get('properties').get('title')
        if page_name == current_page_name:
            page_id = page.get('properties').get('sheetId')
            break
    
    return page_id


def findCreateSpreadSheet(page_name):
    spread_service = spreedSheetService()
    drive_service = driveService()
    sheet_meta = {}

    # find file
    file_list = drive_service.files().list(q="name =  '{}'".format(SHEET_NAME))\
        .execute()
    sheets = file_list.get('files')
    if sheets:
        sheet = sheets[0]
        sheet_meta = spread_service.spreadsheets().get(
                     spreadsheetId=sheet.get('id')).execute()

        # add new sheet if not exist
        if not getPageIdByName(sheet_meta, page_name):
            new_page = {"requests": [{ "addSheet": {"properties":
                                        {"title": page_name}}}]}
            spread_service.spreadsheets().batchUpdate(
                spreadsheetId=sheet.get('id'), body=new_page).execute()

    else:
        # create Drive sheet
        new_sheet = {
            "sheets": [{"properties": {"title": page_name}}],
            "properties": {"title": SHEET_NAME}
        }
        sheet_meta = spread_service.spreadsheets().create(body=new_sheet)\
                     .execute()
    
    return sheet_meta


def calendarService():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    return discovery.build('calendar', 'v3', http=http)


def spreedSheetService():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    return discovery.build('sheets', 'v4', credentials=credentials)


def driveService():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    return discovery.build('drive', 'v3', credentials=credentials)
    

def main():
    # services
    calendar_service = calendarService()
    spread_service = spreedSheetService()
    drive_service = driveService()

    # get dates
    date_now = datetime.datetime.now()
    year = YEAR if YEAR else date_now.year
    month = MONTH if MONTH else date_now.month
    last_day = calendar.monthrange(year, month)[1]

    date_ini = datetime.datetime(year, month, 1, 00).isoformat() + 'Z'
    date_end = datetime.datetime(year, month, last_day, 00).isoformat() + 'Z'
    print('start: ', date_ini)
    print('end: ', date_end)
    page_name = SHEET_PAGE + str(month)

    # get calendar events
    eventsResult = calendar_service.events().list(calendarId='primary',
        timeMin=date_ini, timeMax=date_end, maxResults=300, singleEvents=True,
        orderBy='startTime').execute()
    events = eventsResult.get('items', [])
    event_sort = sorted(events, key=lambda x: x["summary"])
    
    if not events:
        print('No upcoming events found.')

    sheet_meta = findCreateSpreadSheet(page_name)
    sheet_id = sheet_meta.get('spreadsheetId')
    page_id = getPageIdByName(sheet_meta, page_name)

    # array content and head
    array_to_sheet = [['Start', 'End', 'Duration', 'Project', 'Task',
                      'Description ({}-{})'.format(year, month)]]
    # events to array
    for event in event_sort:
        print(event['start'])
        if 'dateTime' in event['start']:
            time_start = event['start'].get('dateTime').encode('utf-8').strip()
            time_end = event['end'].get('dateTime').encode('utf-8').strip()
        if 'date' in event['start']:
            time_start = event['start'].get('date').encode('utf-8').strip()
            time_end = event['end'].get('date').encode('utf-8').strip()

        time_delta = str(dateutil.parser.parse(time_end) \
                    - dateutil.parser.parse(time_start))
        description = event['summary'].encode('utf-8').strip()
        project = extractType(description, PROJECTS)
        type_work = extractType(description, TASKS)

        array_to_sheet.append([prettyDate(time_start), prettyDate(time_end),
                               time_delta, project, type_work, description.decode('utf-8')])

    # clear all sheet
    spread_service.spreadsheets().values().clear(
        spreadsheetId=sheet_id, range=page_name, body={})

    body_val = {"values": array_to_sheet}

    response = spread_service.spreadsheets().values().append(
        spreadsheetId=sheet_id, range=page_name,
        insertDataOption='INSERT_ROWS', valueInputOption='USER_ENTERED',
        body=body_val).execute()

    print(response)
    print("spreadsheet: https://docs.google.com/spreadsheets/d/{}/edit#gid={}"\
          .format(sheet_id, page_id))


if __name__ == '__main__':
    main()


# CREATE NEW SPREAD SHEET PAGEs
