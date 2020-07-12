#Sources:
#https://www.geeksforgeeks.org/reading-excel-file-using-python/
#https://developers.google.com/calendar/v3/reference
from __future__ import print_function
import sys
import os.path
from os import path
import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter import simpledialog
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

    calHolder = []
    #Print List of Calanders
    page_token = None
    while True:
      calendar_list = service.calendarList().list(pageToken=page_token).execute()
      for calendar_list_entry in calendar_list['items']:
        print (calendar_list_entry['summary'])
        calHolder.append({"in": calendar_list_entry['summary'], "cal_id":calendar_list_entry['id']})
      page_token = calendar_list.get('nextPageToken')
      if not page_token:
        break

    cal_msg = "Cals: "
    index = 0
    for dicts in calHolder:
        msg = '\n' + str(index) + ' ' + dicts["in"] + '          '
        cal_msg += msg
        index += 1

    print(cal_msg)
    USER_INP = simpledialog.askinteger(title="Select Cal", prompt=cal_msg)
    print (USER_INP)
    if USER_INP == -1:
        print("it should exit")
        sys.exit(1)
        
    cal_id_inp = ''
    index = 0
    for dicts in calHolder:
        if index == USER_INP:
            cal_id_inp = dicts["cal_id"]
            break
        index += 1
    print(cal_id_inp)

    
    #Adding on sheets service
    sheets_service = build('sheets', 'v4', credentials=creds)

    #Spreadsheet ID
    SPREADSHEET_ID = '15-sqH2xXxN2Oq-VPR-Ei7u9aUIqImjEMFieo32gd1BQ'
    SCHEDULE_SHEET_ID = '1461379716' # 2-Schedule Recording-Instructional Day
    INSTRUCTORS_SHEET_ID = '1867685112' # 1-Approve Courses-Instructors-DropDown Menus
    SAMPLE_RANGE_NAME = '2-Schedule Recording-Instructional Day!A57:Y192'
    INSTRUCTORS_SHEET_RANGE = '1-Approve Courses-Instructors-DropDown Menus!N2:O79'
    STAFF_SHEET_RANGE = '1-Approve Courses-Instructors-DropDown Menus!AG2:AH16'

    # Call the Sheets API
    sheet = sheets_service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])
    print (len(values))

    sheet2 = sheets_service.spreadsheets()
    result2 = sheet2.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=INSTRUCTORS_SHEET_RANGE).execute()
    values2 = result2.get('values', [])
    print (len(values2))

    search = []
    START_DATE = simpledialog.askstring(title="Date From (inclusive)", prompt="Enter the start of the date range (MM/DD/YYYY)" )
    RANGE_START = datetime.datetime.strptime(START_DATE, '%m/%d/%Y')
    END_DATE = simpledialog.askstring(title="Date Until (inclusive)", prompt="Enter the end of the date range (MM/DD/YYYY)" )
    RANGE_END = datetime.datetime.strptime(END_DATE, '%m/%d/%Y')

    if not values:
        print('No data found.')
    else:
        for row in values:
            TEST_DATE = datetime.datetime.strptime(row[0], '%m/%d/%Y')
            if (RANGE_START <= TEST_DATE <= RANGE_END):
                search.append(row)
                print('0 ' + row[0] + ' 1 ' + row[1] + ' 4 ' + row[4] + ' 5 ' + row[5] + ' 6 ' + row[6] + ' 7 ' + row[7] + ' 8 ' + row[8] + ' 9 ' + row[9] + ' 10 ' + row[10] + ' 11 ' + row[11] + ' 12 ' + row[12] + ' 13 ' + row[13] )

    inst_to_email = {}
    if not values2:
        print('No data found.')
    else:
        for row in values2:
            if (len(row) == 2):
                print('0: ' + row[0] + " 1: " + row[1])
                inst_to_email[row[0]] = row[1]
            else:
                print('0: ' + row[0] + " 1: email_not_found@example.com")
                inst_to_email[row[0]] = "email_not_found@example.com"

    staff_to_email = {}
    sheet2 = sheets_service.spreadsheets()
    result2 = sheet2.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=STAFF_SHEET_RANGE).execute()
    values2 = result2.get('values', [])
    print (len(values2))
    if not values2:
        print('No data found.')
    else:
        for row in values2:
            if (len(row) == 2):
                print('0: ' + row[0] + " 1: " + row[1])
                staff_to_email[row[0]] = row[1]
            else if (len(row) == 2):
                print('0: ' + row[0] + " 1: email_not_found@example.com")
                staff_to_email[row[0]] = "email_not_found@example.com"

    for row in search:
        #Get Title/Summary
        summary_in = (row[10] + " - " + row[9])
        if row[11]:
            summary_in += (" - " + row[11] + " MGR ")
        if row[12]:
            summary_in += (" - " + row[12] + " IPS ")
        if row[13]:
            summary_in += (" - " + row[13] + " MA ")

        #Get Location
        loc_in = (row[1])
        if (row[1] == "Chrysler Studio"):
            loc_in = "Chrysler Studio 109 B"
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
        instructor = inst_to_email[row[9]]

        print(instructor)
        list_of_attendees = [
            {'email': instructor}
            ]
        if row[11]:
            a0dee = row[11].split(',')
            a1dee = a0dee[0].split()
            a2dee = a1dee[0]
            print(a2dee)
            list_of_attendees.append({'email': a2dee+'@example.com'})
        if row[12]:
            a0dee = row[12].split(',')
            a1dee = a0dee[0].split()
            a2dee = a1dee[0]
            print(a2dee)
            list_of_attendees.append({'email': a2dee+'@example.com'})
        if row[13]:
            a0dee = row[13].split(',')
            a1dee = a0dee[0].split()
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
        #event = service.events().insert(calendarId='primary', body=event, sendUpdates='all').execute()
        #could possibly make a popup with the HTML link as output
        #print ('Event created: %s' % (event.get('htmlLink')))
        print(event)
    sys.exit(1)

if __name__ == '__main__':
    main()
