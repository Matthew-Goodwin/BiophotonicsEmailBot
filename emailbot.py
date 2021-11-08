from email.message import EmailMessage
import smtplib
import os
from dotenv import load_dotenv
import pandas as pd
import glob as glob
from datetime import datetime
import numpy as np
#Load your email address details - change your details in the env file. Keep the file private
load_dotenv("gmailDetailsPersonal.env")
SENDER = os.environ.get("GMAIL_USER")
PASSWORD = os.environ.get("GMAIL_PASSWORD")

#This function laods all the information from the script section and sends the email.
def send_email(recipient, subject, body):
    msg = EmailMessage()
    msg.set_content(body)
    msg["Subject"] = subject
    msg["From"] = SENDER
    msg["To"] = recipient
    server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
    server.login(SENDER, PASSWORD)
    server.send_message(msg)
    server.quit()


# %% OCT mailing List
fileLocation =r'C:/Users/Matt Goodwin/Documents/Github/BiophotonicsEmailBot//' #input the directory to the excel file with all the meeting details.
#Extract the meeting details for the year
meetingInfo = pd.read_excel(glob.glob(fileLocation + '*.xlsx')[0],sheet_name = "MeetingInfo")
meetingInfo = meetingInfo[meetingInfo['Meeting on?'] == 'Yes']
#Create a list of all the dates
my_dates_list = meetingInfo["Date"].tolist()
#Find out what todays date is
input_date = datetime.now()
#From all the dates in the excel file, keep only those that occur after todays date.
results = [d for d in sorted(my_dates_list) if d > input_date]
#The meeting date for the email is set as the next date a meeting is occuring. 
MeetingDate = results[0] if results else None
#Extract the information from the excelspreadsheet for the next meeting.
MeetingDetails = meetingInfo.loc[meetingInfo[meetingInfo["Date"] == MeetingDate].index[0]]


group=MeetingDetails["Meeting"] #Find out what the meeting is for
groupList = pd.read_excel(glob.glob(fileLocation + '*.xlsx')[0],sheet_name = group) #extract the details from the email list
recipients = groupList["Email"].tolist() 
recipientsName = groupList["Name"]


#Find out who is scheduled to be speaking. Ideally there should be atleast 1 speaker scheduled, if there are two it will include the extra speaker.
speakers = str(MeetingDetails["Speaker1"] + ' and ' + str(MeetingDetails["Speaker2"]))
if str(MeetingDetails["Speaker2"]) == "nan": 
    speakers = str(MeetingDetails["Speaker1"])


#The email template - Change the final line to whoever is in charge.
EmailBody = "Hey team,\n\nWe have a {} meeting Monday ({}/{}) in  {} at {}\nIn the meeting we will have presentations from: {} \n\nAdd any items to the agenda in the google drive\n{} \n\nSee you all then,\nPerson in charge of organising meetings".format(group,MeetingDetails["Date"].day,MeetingDetails["Date"].month, MeetingDetails["Room"], MeetingDetails["Time"].hour, speakers,MeetingDetails["Agenda/Minutes Link"])
#Subject header
subjectHeader = "{} Meeting {}/{}".format(group,MeetingDetails["Date"].day,MeetingDetails["Date"].month)
#Print the email how it will look with all the formatting - any Nans mean something is missing from the template.
print(subjectHeader + '\n\n' + EmailBody)

#Uncomment this final line when you want to actually send the email.
#send_email(recipients, subject=subjectHeader, body=EmailBody)


