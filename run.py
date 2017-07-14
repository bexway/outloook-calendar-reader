#Necessary Installations: pypiwin32, python-dateutil
#SO reference: http://stackoverflow.com/questions/21477599/read-outlook-events-via-python
# https://msdn.microsoft.com/en-us/library/office/ff869026(v=office.15).aspx


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
        date = parse(meetingDate).date()
        time = parse(meetingDate).time()
        # print a.RequiredAttendees
        # print a.OptionalAttendees
        # print a.Organizer
        participants = []
        for r in a.Recipients:
            participants += [str(r)]

        #if the event has appeared before, add to its meeting list
        if subject in appointmentDictionary.keys():
            appointmentDictionary[subject]["Meetings"] += [date.strftime("%m/%d/%Y")]
            appointmentDictionary[subject]["Times"] += [time.strftime("%I:%M %p")]
            appointmentDictionary[subject]["Durations"] += [duration]
            temp = appointmentDictionary[subject]["Participants"]+participants
            appointmentDictionary[subject]["Participants"] = list(set(temp))
        else:
            appointmentDictionary[subject] = {"Subject": subject, "Body": body, "Meetings": [date.strftime("%m/%d/%Y")], "Times":[time.strftime("%I:%M %p")], "Durations": [duration], "Participants":participants}

    resultsfile = open("resultsTally.csv", 'wb')
    fields = ["Subject", "Body", "Number of Occurences", "Date (First)", "Time (First)", "Duration (First)", "Date (Second)", "Time (Second)", "Duration (Second)", "Date (Third)", "Time (Third)", "Duration (Third)", "Further Dates", "Further Times", "Further Durations", "Participants"]
    resultsWriter = csv.DictWriter(resultsfile, fields)
    resultsWriter.writeheader()

    for subject in appointmentDictionary.keys():
        rowDict = {}
        rowDict["Subject"] = appointmentDictionary[subject]["Subject"] if appointmentDictionary[subject]["Subject"] else ""
        rowDict["Body"] = appointmentDictionary[subject]["Body"] if appointmentDictionary[subject]["Body"] else ""
        rowDict["Participants"] = ", ".join(appointmentDictionary[subject]["Participants"]) if appointmentDictionary[subject]["Participants"] else ""
        MeetingWriter(rowDict, appointmentDictionary[subject]["Meetings"], appointmentDictionary[subject]["Times"], appointmentDictionary[subject]["Durations"])
        rowDict["Number of Occurences"] = len(appointmentDictionary[subject]["Meetings"])

        resultsWriter.writerow(rowDict)



def MeetingWriter(rowDict, meetings, times, durations):
    datecount = 0
    for i in range(0, len(meetings)):
        if datecount == 0:
           rowDict["Date (First)"] = meetings[i]
           rowDict["Time (First)"] = times[i]
           rowDict["Duration (First)"] = durations[i]
        elif datecount == 1:
           rowDict["Date (Second)"] = meetings[i]
           rowDict["Time (Second)"] = times[i]
           rowDict["Duration (Second)"] = durations[i]
        elif datecount == 2:
           rowDict["Date (Third)"] = meetings[i]
           rowDict["Time (Third)"] = times[i]
           rowDict["Duration (Third)"] = durations[i]
        else:
            if "Further Dates" in rowDict.keys():
               rowDict["Further Dates"] += ", " + meetings[i]
               rowDict["Further Times"] += ", " + times[i]
               rowDict["Further Durations"] += ", " + durations[i]
            else:
               rowDict["Further Dates"] = meetings[i]
               rowDict["Further Times"] = times[i]
               rowDict["Further Durations"] = durations[i]
        datecount += 1
    return rowDict


if __name__ == "__main__":
    main()
