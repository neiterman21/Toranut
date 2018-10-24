#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
hows basic usage of the Google Calendar API. Creates a Google Calendar API
service object and outputs a list of the next 10 events on the user's calendar.
"""
from __future__ import print_function
from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import datetime
import sys 
import argparse
#reload(sys)
#sys.setdefaultencoding('utf-8')
#hello this is a false line.
#this is demo branch

def parsArgs():
    print ("i am here")
    parser = argparse.ArgumentParser()
    parser.add_argument("--email", help="users google acount mail", action="store")
    parser.add_argument("--שם", help="users name as shown in exle file", action="store")
    parser.add_argument("--xls", help="xls toranut fike", action="store")
    parser.add_argument("--dates", nargs='+' , help="days of toranut", action="store") 
    args = parser.parse_args()
    return args


my_email='neiterman21@gmail.com'
global service


# Setup the Calendar API
def setup_calendar():
    global service
    SCOPES = 'https://www.googleapis.com/auth/calendar'
    store = file.Storage('credentials.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('calendar', 'v3', http=creds.authorize(Http()))
def pars_xls(xls_file):
    return 0

# Call the Calendar API
def Main():
    global service
    args = parsArgs()
    setup_calendar()
#    toranuyot = pars_xls(args.xls)

    myevent = {
    'summary': 'תורנות',
    'location': 'הילל יפה',
    'description': 'הזדמנות לא לישון מחוץ לבית',
    'start': {
        'date': '2018-10-22'
    },
    'end': {
        'date': '2018-10-22'
    },
    'recurrence': [
        'RRULE:FREQ=DAILY;COUNT=1'
    ],
    'attendees': [],
    'reminders': {
        'useDefault': False,
        'overrides': [],
    },
    }
    date = '2018-10-%s'
    for day in args.dates :
        myevent['start']['date'] = (date % day)
        myevent['end']['date']   = (date % day)
        print (myevent['start']['date'])
        service.events().insert(calendarId=args.email , body = myevent).execute()
    '''
    for day in toranuyot:
        if args.name in day['toranim']:
            myevent['start']['date'] = day['date']
            myevent['end']['date'] = day['date']
            myevent['colorId'] = 'Tomato'
            service.events().insert(calendarId=args.email , body = myevent).execute()
        if args.name in day['vacation']:
            myevent['start']['date'] = day['date']
            myevent['end']['date'] = day['date']
            myevent['colorId'] = 'Tomato'
            service.events().insert(calendarId=args.email , body = myevent).execute()
    '''

if __name__ == '__main__':
    Main()
