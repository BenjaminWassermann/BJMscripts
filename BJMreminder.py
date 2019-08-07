#!/usr/bin/env python3
# BJMreminder.py

# This script checks surgery date/survived
# notifying JEO by email every week
# with a list of upcoming harvest groups

# import statements
import openpyxl, datetime, smtplib, BJMcount
from email.mime.multipart import MIMEMultipart as mm
from email.mime.text import MIMEText as mt

# open botInfo workbook
path = 'botInfo.xlsx'
botInfo = openpyxl.load_workbook(filename = path, data_only = True)
botSheet = botInfo['Login']
contactSheet = botInfo['Contacts']

# get today's date
today = datetime.datetime.today()
strToday = str(today)[:10]

# setup email variables and context
smtp_server = smtplib.SMTP('smtp.gmail.com: 587')
sender_email = botSheet['A2'].value
password = botSheet['B2'].value

# pull all recipients from the list
receiver_email = ''
for row in contactSheet:
    if row[0].value != None:
        receiver_email += row[0].value
        receiver_email += ','
    
msg = mm()
msg['From'] = sender_email
msg['To'] = receiver_email
msg['Subject'] = 'BJM schedule update %s' % strToday
message = 'Good morning,\n\nThis email is to remind you of animals which will need to be processed in the coming weeks.\n\nBelow you will see a list of tat administration dates and animal IDs in that group.\n'

# close book
botInfo.close()

# generate placeholder dictionary for surgery groups
groups = {}

print('Please wait while I work...')

# open BJM and select active sheet
BJMbook = openpyxl.load_workbook(filename = 'BJM.xlsx')
BJMsheet = BJMbook['Data']

print('Book open...')

# iterate through the rows
for row in BJMsheet:

    # checks for valid animal ID number
    if row[0].value not in ['30-Day TRPm2 Study', '45-Minute MCAO/Sham', 'Animal Id', None]:

        # checks if the animal is still alive
        if row[5].value in ['Y', 'y']:

            # get the tat admin date
            tatDate = row[6].value

            # check if tatDate has passed
            if today < tatDate:

                # get the weekday number
                wd = datetime.date.weekday(tatDate)

                # if not a Monday
                if wd != 0:

                    # move tat date back to monday
                    tatDate = tatDate - datetime.timedelta(days=wd)

                # get a string from surgeryDate
                strTatDate = str(tatDate)[:10]

                # checks if strSurgeryDate is already a key in groups
                if strTatDate not in groups.keys():

                    # if not already in groups, adds with an empty list
                    groups.update({strTatDate : []})

                # add animal ID to group
                groups[strTatDate].append(row[0].value)

print('Scan complete!')
print('Preparing tatSchedule and email message...')

# open text file
file = open('tatSchedule.txt', 'w')

# harvest schedule instructions
instructions = '\n     Monday - tat-M2NX\n     Tuesday - Behavior\n     Wednesday - Harvest\n     Thursday - Cryoprotection\n\n\n'

# write instructions to message
message += instructions
# iterate through keys
for key, value in groups.items():

    header = 'Administer tat-M2NX on %s\nto the following animals:\n' % key

    # print the group once
    message += header
    
    # iterate through list of IDs
    for animal in value:

        idIndent = '     %s\n' % animal

        # print the animal id
        message += idIndent

    # add an extra space after each group
    message += '\n'

# add countMessage to the message
message += BJMcount.countMessage
file.write(message)

# add a signature to the message
message += 'Thanks,\nYour friendly neighborhood bot\n\nDo not reply to this email - I am a bot\nSend any replies to Benjamin.Wassermann@cuanschutz.edu'

# close files
BJMbook.close()
file.close()

print('tatSchedule and message complete!')
print('Please wait while I send an email to %s...' % receiver_email)

# add attachment to message
file = open('tatSchedule.txt', 'r')
attachment = mt(file.read())
attachment.add_header('Content-Disposition', 'attachment', filename='tatSchedule.txt.')
msg.attach(attachment)

# add message to body
msg.attach(mt(message, 'plain'))

# ping the server
smtp_server.starttls()

# login to the server
smtp_server.login(msg['From'], password)

# send the message from the server
smtp_server.sendmail(msg['From'], msg['To'].split(','), msg.as_string())

# quit server and close files
smtp_server.quit()
file.close()

print('Email sent!')
