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
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import xlrd 

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']
root = tk.Tk()
root.withdraw()

if (not (path.exists("token.pickle"))):
    tkinter.messagebox.showinfo( "Excel to Google Event", "You will be prompted to login & give permission to Google Cal")
    
filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
print(filename)
