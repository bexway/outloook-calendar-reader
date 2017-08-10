# Outlook Calendar Reader
A Python script that uses [pywin32](https://pypi.python.org/pypi/pypiwin32) to read a Microsoft Outlook Calendar and build a csv file of the scheduled appointments within a range of dates specified by the user.

Events of the same subject line are grouped into one csv line, which accumulates all of the meeting dates. Recurring events will be included and accumulated in this format, as well as events that simply have the same subject name.

## Installation and Running
1. If you haven't already, download and install [Python 2.7.x](https://www.python.org/) and ensure it's in your system path. To check whether this has all been done, go to your command line/terminal and use the `python` command. If it begins running the Python shell, your installation was successful.

2. Use a package installer (recent Python versions come with pip) to install the following packages:

    1. [pypiwin32](https://pypi.python.org/pypi/pypiwin32)
    2. [python-dateutil](https://dateutil.readthedocs.io/en/stable/)

3. Download the `run.py` file from this repository and set your command line/terminal's working directory to the folder containing it. (If there is already a file named `resultsTally.csv` in the folder with `run.py` and you would like to keep it, move it elsewhere. Otherwise, when the script runs, it will overwrite that file.)

4. Ensure you are logged in to the account whose appointments should be tallied, then run the Python script. (It may take a few seconds to start if Outlook isn't already open, since it has to run the program to access the Calendar information.) When prompted, specify the day to start tallying appointments and the day to end. All appointments occuring on or between those days will be counted.

5. When the script has finished, a file named `resultsTally.csv` should be written in the same directory as run.py.

## How it Works

The script uses the win32com client (from the pypiwin32 package linked above) to launch Outlook and read events from the Calendar of the signed-in Account. The events are sorted by start time, then restricted to a list containing only the appointments from the user-input time range.

For each appointment, it collects the date, time, duration and recipients of the event to store in the tally. If the event subject matches a subject name already tallied, it won't create a new entry, but will add the date, time, and duration from the new instance to the existing tally for that subject.

The first three instances of an event's dates, times, and durations are written to the csv into separate columns, and the others are combined into one more column.

## Goals

I would like to allow the user to specify the directory and filename for the resulting csv. I would also like to work on parsing the Invitees list for events to see who else is involved in events.