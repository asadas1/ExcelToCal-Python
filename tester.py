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

dict_of_locations = {'Chrysler 133':0, 'Chrysler 151':1, 'Chrysler 165':2, 'Chrysler Studio':3}
list_of_variables = []

def main():
    master = Tk()
    for i in range(len(dict_of_locations)):
        list_of_variables.append(IntVar(master))
    tk.Label(master, text="This is the header:", padx = 10, pady = 5, anchor = 'center').grid(row=0)
    endrow = 0
    for i, location in enumerate(dict_of_locations):
        tk.Label(master, text=location, padx = 10, pady = 5).grid(row=i+1)
        ee = tk.Entry(master, textvariable=list_of_variables[i])
        ee.delete(0, END)
        ee.insert(0, str(dict_of_locations[location]))
        ee.grid(row =i+1, column = 1)
        endrow = i+1

    endrow += 2
    
    def callback():
        for variable in list_of_variables:
            print(variable.get())
        master.destroy
        sys.exit(1)
    
    b = Button(master, text="get", width=10, command=callback)
    b.grid(row=endrow+2, column=0)

    master.mainloop()

    

if __name__ == '__main__':
    main()
