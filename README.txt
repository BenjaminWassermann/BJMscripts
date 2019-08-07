README for BJM scripts

1. BJM.xlsx
	
	Fill in animal IDs in the first row, please do not use periods ('.') or forward/back slashes ('\ /')
	Check that tat admin date formula is correctly calculating the date and fix it if not.
	Several rows use dropdowns, please do not stray from these. Use M or F for sex, Stroke or Sham for
	procedure, and Y or N for survived.

2. botInfo.xlsx

	A source of email login info and email contacts list. Before use, update this spreadsheet with
	bot email login info and email contacts list. Currently contains only sample data.

3. tatSchedule.txt

	Placeholder text document to be filled with the details of animal processing dates and 
	m/f stroke/sham counts. 

4. BJMcount.py

	Script reads in BJM.xlsx, checking the first cell in each row to check for legally formed
	animal IDs. If a row features an appropriate animal ID, the script will read in sex,
	procedure, and survival status. Counts each male or female stroke or sham and stores it
	on the count sheet of BJM.xlsx. Also stores the information as a formatted string in countMessage

5. BJMreminder.py

	Script reads in email sender info and recipient list from botInfo.xlsx.

	Imports BJMcount to use BJMcount.countMessage string. This allows the formatted count string
	to be included in the final message in addition to on the spreadsheet.

	Script reads in BJM.xlsx, checking the first cell in each row to check for legally formed
	animal IDs. If a row features an appropriate animal ID, the script will read in the tat
	administration date. If this date is not on a Monday, date is recalculated to Monday of
	the same week. If this date has not past, keys in the dictionary groups are queried for 
	matching tat admin dates. If not found, adds a dictionary entry with key = string of tat
	admin date and value = [] (an empty list). Then the animal ID is appended to the list 
	found at dictionary{stringTatDate}.

	For each key in groups, a header is added to the ongoing message with the tat admin date.
	Below each date header, each item in the list at that dictionary key are printed to the
	message. The message is recorded to tatSchedule.txt before an email signature is appended.

	Finally, the message is copied into the body of an email addressed to the full list of
	recipients. The tatSchedule text file is also updated and attached to the email.

6. BJMreminder.bat

	A simple batch file which references the Python interpreter and the location of BJMreminder.py.
	This batch file allows me to set up a task in Windows Task Scheduler so that this code
	is executed every Monday at 10am.