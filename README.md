# BiophotonicsEmailBot
This repository is to hold the simple code to send an email to all group members of the upcoming meeting.

The code contains
1) emailbot.py - python script to send the emails
2) BiophotonicsMeetingMaster.xlsx - Excel spreadsheet holding the information about the meeting schedule.
3) gmailDetailsPersonal - env file for the user to store their email details as an environment variable.


The user will basically have to setup the spreadsheet with all the meeting details. The script will pull all those details and format an email to send which informs the group memebers of the upcoming meeting.

Once you have the file there are a few things you HAVE to do:
1) add your email login details to the env file - I couldn't get this working with the university domain email addresses and instead you need to use a personal account as you need to change a few settings to allow python to access your email. Go here and allow less secure apps (set to ON)
https://myaccount.google.com/lesssecureapps

2) In the emailbot.py script, adjust the body of the text to include your name at the bottom of the email (see the end of the string). Also adjust the fileLocation variable to match the directory where the spreadsheet is stored.

3) In the excel spreadsheet, make sure the date and meeting details in the line are filled in correctly. Under the heading "Meeting" the current options are "Biophotonics Meeting" or "OCT Meeting". Those two variables match two tabs in the excel spreadsheet, those respective tabs hold the recipient information, i.e. everyone who it will send an email to. Adjust those respective lists to include the correct people.

If the meeting changes, then you can create a new tab and new list of recipients just make sure the name of the tab matches the name under the "meeting" heading.

Any problems, email: Matthew.Scott.Goodwin@gmail.com

