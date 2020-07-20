#Sources:
#https://www.geeksforgeeks.org/reading-excel-file-using-python/
#https://developers.google.com/calendar/v3/reference
from __future__ import print_function
import sys
import logging
import os.path
from os import path
import gspread
import tkinter as tk
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter import simpledialog
import tkinter.messagebox
import datetime
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import xlrd

#This sets up logging for exceptions and output basically
log = open("log.txt", "a")
#sys.stdout = log

LOG_FILENAME = 'exceptions.txt'
logging.basicConfig(filename=LOG_FILENAME, level=logging.DEBUG)
def my_handler(type, value, tb):
    logging.exception("Uncaught exception: {0}".format(str(value)))
#sys.excepthook = my_handler

# If modifying these scopes, delete the file token.pickle. "spreadsheets.google.com/feeds"
SCOPES = ['https://www.googleapis.com/auth/calendar', 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive"]

#Root window for TK
root = tk.Tk()
root.withdraw()

#Spreadsheet ID's and various other information.
#Can be re-written to access via INI?
    
SPREADSHEET_ID = '15-sqH2xXxN2Oq-VPR-Ei7u9aUIqImjEMFieo32gd1BQ'
SCHEDULE_SHEET_ID = '1461379716' # 2-Schedule Recording-Instructional Day
INSTRUCTORS_SHEET_ID = '1867685112' # 1-Approve Courses-Instructors-DropDown Menus
SAMPLE_RANGE_NAME = '2-Schedule Recording-Instructional Day!A57:AA'

# Range for Sheet 1- is N to AN
INSTRUCTORS_SHEET_RANGE = '1-Approve Courses-Instructors-DropDown Menus!N2:AN799'
STAFF_SHEET_RANGE = '1-Approve Courses-Instructors-DropDown Menus!AG2:AH16'
LOCATION_SHEET_RANGE = '1-Approve Courses-Instructors-DropDown Menus!R2:R12'




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

    #Build gspread
    client = gspread.authorize(creds)

    #Open gspread
    gsheet = client.open("Nexus Recording Schedule - Master")
    gworksheet = gsheet.worksheet("2-Schedule Recording-Instructional Day")

    #Init & print list of Cals
    calHolder = []
    page_token = None
    while True:
      calendar_list = service.calendarList().list(pageToken=page_token).execute()
      for calendar_list_entry in calendar_list['items']:
        print (calendar_list_entry['summary'])
        calHolder.append({"in": calendar_list_entry['summary'], "cal_id":calendar_list_entry['id']})
      page_token = calendar_list.get('nextPageToken')
      if not page_token:
        break

    #Append to single string in order to display in msgbox
    cal_msg = "Please match the numbers for the Calendars for each location, if no match exists enter \'-1\' " + '\n' +  "Calanders on your account: " + '\n'
    cal_msg2 = '\n'
    index = 0
    for dicts in calHolder:
        msg = '\n' + ' [ ' + str(index) + ' ]:   ' + dicts["in"] + '          '
        cal_msg2 += msg
        index += 1

    #Adding on sheets service
    sheets_service = build('sheets', 'v4', credentials=creds)



    
    #Prompt user for selection via messagebox
    cal_msg2 += '\n'
    print(cal_msg)
    print(cal_msg2)
    USER_INP = -1

    #Yes this is janky/bad code, but I mean cmon
    sheet3 = sheets_service.spreadsheets()
    result3 = sheet3.values().get(spreadsheetId=SPREADSHEET_ID, range=INSTRUCTORS_SHEET_RANGE).execute()
    values3 = result3.get('values', [])
    for row in values3:
        if (len(row) >= 5):
            print(row[4])
        if (len(row) >= 16):
            print("    MGR: " + row[15])
        if (len(row) >= 17):
            print("    IPS: " + row[16])
        if (len(row) >= 18):
            print("    MA: " + row[17])
        if (len(row) >= 19):
            print("    Other: " + row[18])
        if (len(row) >= 21):
            print("         " + row[19] + ": " + row[20])
        if (len(row) >= 22):
            print("    CCBK: " + row[21])
        if (len(row) >= 23):
            print("    o-CCBK: " + row[22])
        if (len(row) >= 25):
            print("    nc-mary: " + row[24])
        if (len(row) >= 26):
            print("    nc-dru: " + row[25])
        if (len(row) >= 27):
            print("    nc-paul: " + row[26])

    dict_of_locations = {}
    if path.exists("calconfig.pickle"):
        with open('calconfig.pickle', 'rb') as handle:
            dict_of_locations = pickle.load(handle)
    else:
        if not values3:
            print('No data found.')
        else:
            for row in values3:
                if (len(row) >= 5):
                    if row[4]:
                        dict_of_locations[row[4]] = -1
        with open('calconfig.pickle', 'wb') as handle:
            pickle.dump(dict_of_locations, handle, protocol=pickle.HIGHEST_PROTOCOL)
    
    list_of_variables = []

    master = Tk()

    def callback():
        for i, location in enumerate(dict_of_locations):
            print(list_of_variables[i].get())
            dict_of_locations[location] = list_of_variables[i].get()
        master.destroy()
        master.quit()
        print(dict_of_locations)
        with open('calconfig.pickle', 'wb') as handle:
            pickle.dump(dict_of_locations, handle, protocol=pickle.HIGHEST_PROTOCOL)

    master.protocol("WM_DELETE_WINDOW", callback)
    for i in range(len(dict_of_locations)):
        list_of_variables.append(IntVar(master))
        
    tk.Label(master, text=cal_msg, padx = 10, pady = 5, anchor = 'center').grid(row=0)
    tk.Label(master, text=cal_msg2, padx = 10, pady = 5, justify = 'left').grid(row=1)
    endrow = 0
    for i, location in enumerate(dict_of_locations):
        tk.Label(master, text=location, padx = 10, pady = 5).grid(row=i+2)
        ee = tk.Entry(master, textvariable=list_of_variables[i])
        ee.delete(0, END)
        ee.insert(0, dict_of_locations[location])
        ee.grid(row = i+2, column = 1)
        endrow = i+2

    endrow += 2

    
    b = Button(master, text="Submit", width=10, command=callback)
    b.grid(row=endrow+2, column=0)
    master.mainloop()

    #Convert indexes to CAL ID
    for i, location in enumerate(dict_of_locations):
        for j, cal in enumerate(calHolder):
            if (dict_of_locations[location] == j):
                dict_of_locations[location] = cal["cal_id"]
                
    print(dict_of_locations)



    #Get what kind of method to select events
    search_method = 0
    window = Tk()
    v = IntVar(window)
    v.set(0)

    def ShowChoice():
        print(v.get())
        if (v.get() == 0):
            sys.exit(1)
        search_method = v.get()
        window.destroy()
        window.quit()

    tk.Label(window, 
             text="""Choose method for selecting events:""",
             padx = 20, pady = 5).pack()
    tk.Radiobutton(window, 
                  text="By Date RANGE (MM/DD/YYYY)",
                  indicatoron = 0,
                  width = 20,
                  padx = 20, 
                  variable=v, 
                  command=ShowChoice,
                  value=1).pack()
    tk.Radiobutton(window, 
                  text="By Row In RANGE ex: (1-9999)",
                  indicatoron = 0,
                  width = 20,
                  padx = 20, 
                  variable=v, 
                  command=ShowChoice,
                  value=2).pack()
    tk.Radiobutton(window, 
                  text="By Row in LIST ex: (64, 65, 77, 81)",
                  indicatoron = 0,
                  width = 20,
                  padx = 20, 
                  variable=v, 
                  command=ShowChoice,
                  value=3).pack()
        
    window.mainloop()

    search_method = v.get()
    print(search_method)
    #sys.exit(1)

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

    #Prompt & convert date range
    search = []
    search_indexes = []
    if search_method == 1:
        START_DATE = simpledialog.askstring(title="Date From (inclusive)", prompt="Enter the start of the date range (MM/DD/YYYY)" )
        RANGE_START = datetime.datetime.strptime(START_DATE, '%m/%d/%Y')
        END_DATE = simpledialog.askstring(title="Date Until (inclusive)", prompt="Enter the end of the date range (MM/DD/YYYY)" )
        RANGE_END = datetime.datetime.strptime(END_DATE, '%m/%d/%Y')
    if search_method == 2:
        START_ROW = simpledialog.askinteger(title="First Row (Inclusive):", prompt="Enter the first row:" )
        END_ROW = simpledialog.askinteger(title="Last Row (Inclusive):", prompt="Enter the Last row:" )
        if (START_ROW > END_ROW):
            print("startstop error 1")
            sys.exit(1)
    if search_method == 3:
        USER_LIST = simpledialog.askstring(title="Enter List of Rows:", prompt="Enter list of rows seperated by Commas. Ex: (16, 22, 2, 1999)" )
        ROW_LIST = USER_LIST.split(",")

    #Search for valid entries within range
    s_index = 0
    if not values:
        print('No data found.')
    else:
        for row in values:
            if not (row[0]):
                continue
            if search_method == 1:
                TEST_DATE = datetime.datetime.strptime(row[0], '%m/%d/%Y')
                if (RANGE_START <= TEST_DATE <= RANGE_END):
                    search.append(row)
                    search_indexes.append(s_index)
                    print('0 ' + row[0] + ' 1 ' + row[1] + ' 4 ' + row[4] + ' 5 ' + row[5] + ' 6 ' + row[6] + ' 7 ' + row[7] + ' 8 ' + row[8] + ' 9 ' + row[9] + ' 10 ' + row[10] + ' 11 ' + row[11] + ' 12 ' + row[12] + ' 13 ' + row[13] + ' 18 ' + row[18] + ' 19 ' + row[19] + ' 25 ' + row[25] + ' 26 ' + row[26])
            if search_method == 2:
                if (START_ROW <= int(row[26]) <= END_ROW):
                    search.append(row)
                    search_indexes.append(s_index)
                    print('0 ' + row[0] + ' 1 ' + row[1] + ' 4 ' + row[4] + ' 5 ' + row[5] + ' 6 ' + row[6] + ' 7 ' + row[7] + ' 8 ' + row[8] + ' 9 ' + row[9] + ' 10 ' + row[10] + ' 11 ' + row[11] + ' 12 ' + row[12] + ' 13 ' + row[13] + ' 18 ' + row[18]  + ' 19 ' + row[19] + ' 25 ' + row[25] + ' 26 ' + row[26])
            if search_method == 3:
                for rowval in ROW_LIST:
                    if (int(rowval) == int(row[26])):
                        search.append(row)
                        search_indexes.append(s_index)
                        print('0 ' + row[0] + ' 1 ' + row[1] + ' 4 ' + row[4] + ' 5 ' + row[5] + ' 6 ' + row[6] + ' 7 ' + row[7] + ' 8 ' + row[8] + ' 9 ' + row[9] + ' 10 ' + row[10] + ' 11 ' + row[11] + ' 12 ' + row[12] + ' 13 ' + row[13] + ' 18 ' + row[18]  + ' 19 ' + row[19] + ' 25 ' + row[25] + ' 26 ' + row[26])
            s_index += 1

    #Read in instructor emails
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

    #Read in staff emails
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
            if (len(row) == 1):
                print('0: ' + row[0] + " 1: email_not_found@example.com")
                staff_to_email[row[0]] = "email_not_found@example.com"

    #Setup list of events for printing
    event_printlist = []

    #Begin creating & sending events
    s_index = 0
    for row in search:
        #skip if the event was already made
        if row[25] != 'N':
            print("skipped " + row[10] + " " + row[0])
            continue
        #Convert location to CAL ID
        if dict_of_locations[row[1]] == -1:
            print("skipped " + row[10] + " " + row[0] + " No Cal")
            continue
        
        gworksheet.update_cell(int(row[26]), 26, "Y")
        s_index += 1
        
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

        #Get Attendees 
        #List of attendees is a "list of dicts" which is the input the JSON object "event" wants
        instructor = inst_to_email[row[9]]
        print(instructor)

        #Staff
        staff_holder = ""
        list_of_attendees = [
            {'email': instructor}
            ]
        if row[11]:
            their_email = staff_to_email[row[11]]
            list_of_attendees.append({'email': their_email})
            staff_holder = row[11]
            print(their_email)
        if row[12]:
            their_email = staff_to_email[row[12]]
            list_of_attendees.append({'email': their_email})
            staff_holder = row[12]
            print(their_email)
        if row[13]:
            their_email = staff_to_email[row[13]]
            list_of_attendees.append({'email': their_email})
            staff_holder = row[13]
            print(their_email)

        #Credit/Noncredit list(?) WIP
        print(row[18])
        if (row[18] == "Credit"):
            print("It's a Credit Course. ")
            print(staff_holder)
            if (staff_holder == "Brandon"):
                list_of_attendees.append({'email': "reillym@umich.edu", 'optional': 1})
                list_of_attendees.append({'email': "skash@umich.edu", 'optional': 1})
            if (staff_holder == "Mary"):
                list_of_attendees.append({'email': "bsandusk@umich.edu", 'optional': 1})
                list_of_attendees.append({'email': "skash@umich.edu", 'optional': 1})



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

        print(dict_of_locations[row[1]])
        #Uses the service to insert the event
        #event = service.events().insert(calendarId=dict_of_locations[row[1]], body=event, sendUpdates='all').execute()
        #could possibly make a popup with the HTML link as output
        #print ('Event created: %s' % (event.get('htmlLink')))
        #event_link = event.get('htmlLink')
        event_link = "google.com"
        event_printlist.append({'summary':summary_in, 'date':row[0], 'link':event_link})
        print(event)


    f = open("CreatedEvents.html", 'w')
    f.write("<h1>Created the Following Events:</h1>" + '\n' + "<blockquote>")
    for event in event_printlist:
        f.write('\n' + "<p>" + event['summary'] + ' ' + event['date'] + ':' + "</p>")
        f.write('\n' + "<p><a href=\"" + event['link'] + "\">" + event['link'] + "</a></p>")
    f.write('\n' + "</blockquote>")
    f.close()
    os.startfile("CreatedEvents.html")
    sys.exit(1)

if __name__ == '__main__':
    main()
