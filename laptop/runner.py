#Sources:
#https://www.geeksforgeeks.org/reading-excel-file-using-python/
#https://developers.google.com/calendar/v3/reference

from __future__ import print_function
import os.path
from os import path
import tkinter as tk
from tkinter.filedialog import askopenfilename
import tkinter.messagebox
import datetime
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import xlrd 

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar', 'https://www.googleapis.com/auth/spreadsheets']

#Root window for TK
root = tk.Tk()
root.withdraw()

def main():
    # A quick check to see if the token already exists.
    if (not (path.exists("token.pickle"))):
        tkinter.messagebox.showinfo( "Excel to Google Event", "You will be prompted to login & give permission to Google Cal")
    
    #This is taken directly from the Google API Quickstart guide
    """Shows basic usage of the Google Calendar API.
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
    
    #Here the service is built with credentials & we can move on to creating the event
    service = build('calendar', 'v3', credentials=creds)

    #Adding on sheets service
    sheets_service = build('sheets', 'v4', credentials=creds)

    #Spreadsheet ID
    SPREADSHEET_ID = '15-sqH2xXxN2Oq-VPR-Ei7u9aUIqImjEMFieo32gd1BQ'
    SCHEDULE_SHEET_ID = '1461379716' # 2-Schedule Recording-Instructional Day
    INSTRUCTORS_SHEET_ID = '1867685112' # 1-Approve Courses-Instructors-DropDown Menus
    SAMPLE_RANGE_NAME = '2-Schedule Recording-Instructional Day!A57:Y192'

    # Call the Sheets API
    sheet = sheets_service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])
    print (len(values))

    search = []

    if not values:
        print('No data found.')
    else:
        for row in values:
            #print (len(row))
            # Print columns A and E, which correspond to indices 0 and 4.
            if (row[0] == '07/01/2020'):
                search.append(row)
                print('0 ' + row[0] + ' 1 ' + row[1] + ' 4 ' + row[4] + ' 5 ' + row[5] + ' 6 ' + row[6] + ' 7 ' + row[7] + ' 8 ' + row[8] + ' 9 ' + row[9] + ' 10 ' + row[10] + ' 11 ' + row[11] + ' 12 ' + row[12] + ' 13 ' + row[13] )

    for row in search:
        #Get Title/Summary
        summary_in = (row[10])

        #Get Location
        loc_in = (row[1])

        #Get Desc
        desc_in = (row[10])

        #Get Start Time and Date
        start_dts = row[0] + ' ' + row[5]

        #Get End Time and Date
        end_dts = row[0] + ' ' + row[6]

        #Date & timestamp stuff is janky because the JSON object "event" wants RCF formatted time,
        #whereas the Excel file could have any kind of time input, so using strptime with concacted strings is probably the most
        #flexible approach for now
        dto_start = datetime.datetime.strptime(start_dts, '%m/%d/%Y %I:%M %p')
        dto_end = datetime.datetime.strptime(end_dts, '%m/%d/%Y %I:%M %p')

        #Get Attendees // currently not implemented
        #List of attendees is a "list of dicts" which is the input the JSON object "event" wants
        #attendees = ["lpage@example.com", "ddage@example.com"]
        attendee = row[9].split(',')
        sttendee = attendee[0]

        print(sttendee)
        list_of_attendees = [
            {'email': sttendee+'@example.com'}
            ]
        if row[11]:
            print (row[11])
            a0dee = row[11].split(',')
            a1dee = a0dee[0].split()
            if a1dee:
                a2dee = a1dee[0]
                print(a2dee)
                list_of_attendees.append({'email': a2dee+'@example.com'})
        if row[12]:
            print (row[12])
            a0dee = row[12].split(',')
            print(a0dee)
            a1dee = a0dee[0].split()
            print(a1dee)
            if a1dee:
                a2dee = a1dee[0]
                print(a2dee)
                list_of_attendees.append({'email': a2dee+'@example.com'})
        if row[13]:
            print (row[13])
            a0dee = row[11].split(',')
            a1dee = a0dee[0].split()
            if a1dee:
                a2dee = a1dee[0]
                print(a2dee)
                list_of_attendees.append({'email': a2dee+'@example.com'})
        #Is a WIP


        #The actual JSON style event object, time zone is static just because not really necessary 
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
          'attendees': list_of_attendees,
          'reminders': {
            'useDefault': False,
            'overrides': [
              {'method': 'email', 'minutes': 24 * 60},
              {'method': 'popup', 'minutes': 10},
            ],
          },
        }
        
        #Uses the service to insert the event
        event = service.events().insert(calendarId='primary', body=event, sendUpdates='all').execute()
        #could possibly make a popup with the HTML link as output
        print ('Event created: %s' % (event.get('htmlLink')))

if __name__ == '__main__':
    main()
