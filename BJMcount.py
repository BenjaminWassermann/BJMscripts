#!/usr/bin/env python3
# BJMcount.py

# This script checks BJM and counts the number of survived
# male and female strokes and shams

import openpyxl

print('Please wait while I count the sheet...')

# open book and select data and count sheets
book = openpyxl.load_workbook('BJM.xlsx')
dataSheet = book['Data']
countSheet = book['Counts']

# generate placeholders for mouse counts
mStroke = 0
mSham = 0
fStroke = 0
fSham = 0

# iterate through the rows containing animal IDs
for row in dataSheet:

    # checks for valid animal ID number
    if row[0].value not in ['30-Day TRPm2 Study', '45-Minute MCAO/Sham', 'Animal Id', None]:

        # checks if the animal is still alive
        if row[5].value in ['Y', 'y']:

            # check if male
            if row[1].value in ['M', 'm']:

                # checks if procedure is stroke
                if row[3].value in ['Stroke', 'stroke', 'STROKE']:

                    # add 1 to mStroke
                    mStroke += 1

                # checks if procedure is sham
                elif row[3].value in ['Sham', 'sham', 'SHAM']:

                    # add 1 to mSham
                    mSham += 1

                # improperly coded
                else:
                    print('%s procedure must be stroke or sham...' % row[0].value)

                
            # checks if female
            elif row[1].value in ['F', 'f']:

                # checks if procedure is stroke
                if row[3].value in ['Stroke', 'stroke', 'STROKE']:

                    # add 1 to fStroke
                    fStroke += 1

                # checks if procedure is sham
                elif row[3].value in ['Sham', 'sham', 'SHAM']:

                    # add 1 to fSham
                    fSham += 1

                # improperly coded
                else:
                    print('%s procedure must be stroke or sham...' % row[0].value)

            # improperly coded
            else:
                print('%s sex must be M or F...' % row[0].value)

# populate countSheet with counts
countSheet['A2'].value = mStroke
countSheet['B2'].value = mSham
countSheet['A4'].value = fStroke
countSheet['B4'].value = fSham

countMessage = ('Animals survived so far:\nMale stroke: %s\nMale sham: %s\nFemale stroke: %s\nFemale sham: %s\n\n' % (mStroke, mSham, fStroke, fSham))

print('Counts complete and BJM.xlsx updated!')
print(countMessage)

# save and close workbook
book.save('BJM.xlsx')
book.close()
