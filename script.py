'''
    Python code to send email to a list of emails from a spreadsheet
'''
# import the required libraries
# from cgi import test
import pandas as pd
from email import message
import smtplib 
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


# your email details

SERVER = 'smtp.gmail.com'  # your smtp server
PORT  = 587    # your port number
FROM  =  'tonychelseafan@gmail.com'    # your from email id
PASS  = 'hsghdhhd89nssbs' # your email id password

# Authentication part
server = smtplib.SMTP(SERVER,PORT)

server.set_debuglevel(1) # set 1 to help in debugging
server.ehlo()
server.starttls() # start TLS connection which is secure connection

server.login(FROM,PASS)


  
# reading the spreadsheet
email_list = pd.read_excel('data/data.xlsx')


print('Getting the names and the emails................')
# getting the names and the emails
names = email_list['Full name']
emails = email_list['Email Address']
print(emails)

# email composing
print('Composing Email................')
# creating a message body
msg = MIMEMultipart()

msg['Subject'] = 'Invitation:: FAMILY NOVEMBER MEETING AT 4pm'
msg['From'] = FROM

print('email body settings...............')

# iterate through the records
for i in range(len(emails)):
  
    # for every record get the name and the email addresses
    name = names[i]
    email = emails[i]
  
    # the message to be emailed
    message = 'Greetings,  \n \nA kind remeinder to join our November family meeting. You can find information about this meeting below.\n \nTopic: FAMILY NOVEMBER MEETING \nTime: Feb 01 (Tuesday), 2022 \n \nJoin Zoom Meeting \nhttps://zoom.us/j/98889865931?pwd=SmJJfhrgyfghdshgdf \n\nMeeting ID: 678 867 900 \nPasscode: 0000001  \n\nLooking forward to see you today, cheers \n\nRegards'
    msg.attach(MIMEText(message, 'plain'))

    # sending the email
    server.sendmail(FROM, [email], msg.as_string())

print('Email Sent .....') 

# close the smtp server
server.close()


