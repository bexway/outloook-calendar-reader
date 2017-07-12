#Necessary Installations: pypiwin32, python-dateutil
#SO reference: http://stackoverflow.com/questions/21477599/read-outlook-events-via-python
import win32com.client, datetime
from dateutil.parser import *
# from dateutil.tz import *
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
        subject = str(a.Subject)
        body = str(a.Body.encode("utf8"))
        duration = str(a.duration)
        print subject
        # print body
        # print parse(meetingDate).time().strftime("%I:%M %p")
        # print parse(meetingDate).date().strftime("%m/%d/%Y")
        date = parse(meetingDate).date()
        time = parse(meetingDate).time()
        # print a.RequiredAttendees
        # print a.OptionalAttendees
        # print a.Organizer
        # print a.Recipients.Item(1)


        #if the event has appeared before, add to its meeting list
        if subject in appointmentDictionary.keys():
            appointmentDictionary[subject]["Meetings"] += [date.strftime("%m/%d/%Y")]
        else:
            appointmentDictionary[subject] = {"Subject": subject, "Body": body, "Meetings": [date.strftime("%m/%d/%Y")], "Time":time.strftime("%I:%M %p"), "Duration": duration, "Participants":[]}

    resultsfile = open("resultsTally.csv", 'wb')
    fields = ["Subject", "Body", "Number of Occurences", "Date (First)", "Date (Second)", "Date (Third)", "Date (Fourth)", "Date (Fifth)", "Date (Sixth)", "Further Dates", "Time", "Duration", "Participants"]
    resultsWriter = csv.DictWriter(resultsfile, fields)
    resultsWriter.writeheader()

    for subject in appointmentDictionary.keys():
        rowDict = {}
        rowDict["Subject"] = appointmentDictionary[subject]["Subject"] if appointmentDictionary[subject]["Subject"] else ""
        rowDict["Body"] = appointmentDictionary[subject]["Body"] if appointmentDictionary[subject]["Body"] else ""
        rowDict["Time"] = appointmentDictionary[subject]["Time"] if appointmentDictionary[subject]["Time"] else ""
        rowDict["Duration"] = appointmentDictionary[subject]["Duration"] if appointmentDictionary[subject]["Duration"] else ""
        rowDict["Participants"] = appointmentDictionary[subject]["Participants"] if appointmentDictionary[subject]["Participants"] else ""
        MeetingWriter(rowDict, appointmentDictionary[subject]["Meetings"])
        rowDict["Number of Occurences"] = len(appointmentDictionary[subject]["Meetings"])

        resultsWriter.writerow(rowDict)



def MeetingWriter(rowDict, meetings):
    datecount = 0
    for date in meetings:
        if datecount == 0:
           rowDict["Date (First)"] = date
        elif datecount == 1:
           rowDict["Date (Second)"] = date
        elif datecount == 2:
           rowDict["Date (Third)"] = date
        elif datecount == 3:
           rowDict["Date (Fourth)"] = date
        elif datecount == 4:
           rowDict["Date (Fifth)"] = date
        elif datecount == 5:
           rowDict["Date (Sixth)"] = date
        else:
            if rowDict["Further app.s"]:
               rowDict["Further app.s"] += ", " + date
            else:
               rowDict["Further app.s"] += date
        datecount += 1
    return rowDict


if __name__ == "__main__":
    main()
