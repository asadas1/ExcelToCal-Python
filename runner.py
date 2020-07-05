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
SCOPES = ['https://www.googleapis.com/auth/calendar.events']

#Root window for TK
root = tk.Tk()
root.withdraw()

# Give the location of the file 
loc = askopenfilename(title = "Select EXCEL file",filetypes = (("xlsx files","*.xlsx"), ("all files","*.*")) )
  
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

dto_start = datetime.datetime.strptime(start_dts, '%m-%d-%Y %I:%M %p')
dto_end = datetime.datetime.strptime(end_dts, '%m-%d-%Y %I:%M %p')

#Get Attendees
#attendee = (sheet.cell_value(7,1))
attendees = ["lpage@example.com", "ddage@example.com"]
list_of_attendees = [
    {'email': attendees[0] },
    {'email': attendees[1] }
    ]


def main():
    if (not (path.exists("token.pickle"))):
        tkinter.messagebox.showinfo( "Excel to Google Event", "You will be prompted to login & give permission to Google Cal")
    
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

    event = service.events().insert(calendarId='primary', body=event, sendUpdates='all').execute()
    print ('Event created: %s' % (event.get('htmlLink')))

if __name__ == '__main__':
    main()
