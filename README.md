# Outlook Calendar Reader
A Python script that uses [pywin32](https://pypi.python.org/pypi/pypiwin32) to read a Microsoft Outlook Calendar and build a csv file of the scheduled appointments within a range of dates specified by the user.

## Installation
1. If you haven't already, download and install [Python 2.7.x](https://www.python.org/) and ensure it's in your system path. To check whether this has all been done, go to your command line/terminal and use the `python` command. If it begins running the Python shell, your installation was successful.

2. Use a package installer (recent Python versions come with pip) to install the following packages:

    1. [pypiwin32](https://pypi.python.org/pypi/pypiwin32)
    2. [python-dateutil](https://dateutil.readthedocs.io/en/stable/)

3. Download the `run.py` file from this repository and set your command line/terminal's working directory to the folder containing it. (If there is already a file named `resultsTally.csv` in the folder with `run.py` and you would like to keep it, move it elsewhere. Otherwise, when the script runs, it will overwrite that file.)

4. Run the Python script. (It may take a few seconds to start if Outlook isn't already open, since it has to run the program to access the Calendar information.) When prompted, specify the day to start tallying appointments and the day to end. All appointments occuring on or between those days will be counted.

5. When the script has finished, a file named `resultsTally.csv` should be written in the same directory as run.py.