#Necessary Installations: pypiwin32, python-dateutil
#SO reference: http://stackoverflow.com/questions/21477599/read-outlook-events-via-python
import win32com.client, datetime
import dateutil
from dateutil.relativedelta import relativedelta
import re
import csv
import calendar
import Tkinter as tk
import tkFileDialog

def main():
    #Access Outlook and get the events from the calendar
    Outlook = win32com.client.Dispatch("Outlook.Application")
    ns = Outlook.GetNamespace("MAPI")
    appointments = ns.GetDefaultFolder(9).Items

    #Sort the events by occurence and then include recurring events
    appointments.Sort("[Start]")
    appointments.IncludeRecurrences = "True"

    #Generate a dictionary; I need to track appointment dates to count them
    appointmentDictionary = {}
    #Create a regex for time and Subject
    timeregex = re.compile('\d\d/\d\d/\d\d')
    # subjectregex = re.compile("(?P<cancel>\w*cancel(led)?)?[ -,]{,2}(?P<subject>\w.*)[ -,]{,2}(?P<canceltwo>cancel(led)?)?")
    nameregex = re.compile(u'[Nn]ame: ?(?P<name>[\( \)\&;\w]*)', re.UNICODE)
    locationregex = re.compile(u'[Ll]ocation: ?(?P<location>[\( \)\&;\d]*)', re.UNICODE)
    #get names from invitees?
    for a in appointments:
        #grab the date from the meeting time
        meetingDate = str(a.Start)
        subjectMatch = str(a.Subject)
        body = str(a.Body.encode("utf8"))


if __name__ == "__main__":
    main()
