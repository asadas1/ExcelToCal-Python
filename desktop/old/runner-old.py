#Sources:
#https://www.geeksforgeeks.org/reading-excel-file-using-python/
#https://developers.google.com/calendar/v3/reference

from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Reading an excel file using Python 
import xlrd 

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']

# Give the location of the file 
loc = ("Python-Test.xlsx") 
  
# To open Workbook 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 

#Get Title/Summary
summary_in = (sheet.cell_value(1,0))

#Get Location
loc_in = (sheet.cell_value(1,1))

#Get Desc
desc_in = (sheet.cell_value(1,2))

#Get Start Time and Date
starttime_in = (sheet.cell_value(1,3))
startdate_in = (sheet.cell_value(1,4))
start_dts = startdate_in + ' ' + starttime_in

#Get End Time and Date
endtime_in = (sheet.cell_value(1,5))
enddate_in = (sheet.cell_value(1,6))
end_dts = enddate_in + ' ' + endtime_in

dto_start = datetime.datetime.strptime(start_dts, '%m-%d-%Y %H:%M %p')
dto_end = datetime.datetime.strptime(end_dts, '%m-%d-%Y %H:%M %p')

#Get Attendees
#attendee = (sheet.cell_value(7,1))
attendees = ["lpage@example.co", "ddage@example.com"]

def main():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    page_token = None
    while True:
        calendar_list = service.calendarList().list(pageToken=page_token).execute()
        for calendar_list_entry in calendar_list['items']:
            print (calendar_list_entry['summary'])
            print (calendar_list_entry['id'])
        page_token = calendar_list.get('nextPageToken')
        if not page_token:
            break
    
    event = {
      'summary': summary_in,
      'location': loc_in,
      'description': desc_in,
      'start': {
        'dateTime': dto_start.isoformat("T"),
        'timeZone': 'US/Eastern',
      },
      'end': {
        'dateTime': dto_end.isoformat("T"),
        'timeZone': 'US/Eastern',
      },
     # 'recurrence': [
     #   'RRULE:FREQ=DAILY;COUNT=2'
     # ],
      'attendees': [
        {'email': attendees },
      ],
      'reminders': {
        'useDefault': False,
        'overrides': [
          {'method': 'email', 'minutes': 24 * 60},
          {'method': 'popup', 'minutes': 10},
        ],
      },
    }

    event = service.events().insert(calendarId='qs64rv6jvd7lhs7los3r7jh43k@group.calendar.google.com', body=event, sendUpdates='all').execute()
    print ('Event created: %s' % (event.get('htmlLink')))


    # Call the Calendar API
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    print('Getting the upcoming 10 events')
    events_result = service.events().list(calendarId='qs64rv6jvd7lhs7los3r7jh43k@group.calendar.google.com', timeMin=now,
                                        maxResults=10, singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        print('No upcoming events found.')
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        print(start, event['summary'])


if __name__ == '__main__':
    main()
