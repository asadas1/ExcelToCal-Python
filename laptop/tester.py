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

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar', 'https://www.googleapis.com/auth/spreadsheets', "https://www.googleapis.com/auth/drive", "spreadsheets.google.com/feeds"]

#Root window for TK
root = tk.Tk()
root.withdraw()



def main():
    master = Tk()
    
    tk.Label(master, text="This is the header:", padx = 10, pady = 5, anchor = 'center').grid(row=0)
    tk.Label(master, text="First Name", padx = 10, pady = 5).grid(row=1)
    tk.Label(master, text="Last Name", padx = 10, pady = 5).grid(row=2)

    e1 = tk.Entry(master)
    e2 = tk.Entry(master)

    def callback():
        print (e1.get())
        print (e2.get())
        master.destroy
        sys.exit(1)

    e1.grid(row=1, column=1)
    e2.grid(row=2, column=1)
    
    b = Button(master, text="get", width=10, command=callback)
    b.grid(row=3, column=0)

    master.mainloop()

    

if __name__ == '__main__':
    main()
