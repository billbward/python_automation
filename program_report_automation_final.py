""" 
	Import Module Begin
"""

from __future__ import generators  
import pyodbc
import os
import datetime
import json
import smtplib
import time
from time import gmtime, strftime, localtime
from win32com.client import Dispatch
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email import Encoders
import csv
import gspread

""" 
	Import Module End
"""



"""
	Credentials Module Begin
"""

#autoreport credentials and status update toaddrs (bw's phone)
username = 'sender_email'
password = 'sender_pw'
fromaddr = 'sender_email'
toaddrs  = 'my_phone_#'



"""
	Credentials Module End
"""



"""
	Program Report SQL Run and File Creation Module Begin
"""


def ResultIter(cursor, arraysize=1000):
    'An iterator that uses fetchmany to keep memory usage down'
    while True:
        results = cursor.fetchmany(arraysize)
        if not results:
            break
        for result in results:
			yield result


# set up write file variables
root = os.path.join('C:/Users/bgtx.bward/Desktop/BGTX/Reports/Program History Report/scripted_outputs/')
timestamp = datetime.datetime.now().strftime('%Y%m%d')
file_name = 'program_report_dataset_' + timestamp + '.csv'
file = os.path.join(root,file_name)
w = open(file,'w')

# establish connection variables
cnxn = pyodbc.connect("DSN='dsn';UID='user';PW='pw'")
cursor = cnxn.cursor()

# read github voter contact report query - creates a vertica table of the output
file = open('C:/Users/bgtx.bward/Documents/GitHub/bgtx_reports/program history/Program Report Dashboard_Standalone Result Set.sql','r')
program_report_output = file.read()
file.close()

processes = program_report_output

msg = ('\nProgram Toplines: Launched --' + str(datetime.datetime.now().strftime('%Y%m%d %H:%M:%S')) + '\n')

# Sending launch text to bw
server = smtplib.SMTP('smtp.gmail.com:587')
server.starttls()
server.login(username,password)
server.sendmail(fromaddr, toaddrs, msg)
server.quit()

print '\n' + str(datetime.datetime.now().strftime('%Y%m%d %H:%M:%S')) + ':\tgenerating program toplines report dataset'
cursor.execute(processes)

print str(datetime.datetime.now().strftime('%Y%m%d %H:%M:%S')) + ':\t\tgenerating column names'

#get column names and set up as a header row
columns = []

for row in cursor.columns(table='program_toplines_report_today'):
	cleaned_column_name = row.column_name.encode('ascii')
	columns.append(cleaned_column_name)


w.write(','.join(columns) + '\n')

print str(datetime.datetime.now().strftime('%Y%m%d %H:%M:%S')) + ':\t\twriting result set to file '

#pull up that dataset!
cursor.execute('select * from program_toplines_report_today order by 1 desc,2 asc')

for result in ResultIter(cursor):
	cleaned_row = json.dumps(tuple(result)).strip('[]').replace(', ',',')
	w.write(cleaned_row + '\n')
	if result[0]=="State":
		dials = (result[2] + result[5])
		convos = (result[3] + result[6])
		w_num = (result[4] + result[7])
		bw_att = (result[15])
		bw_convos = (result[16])
		door_att = (result[17])
		vdrs = (result[10])
		vr = (result[19])
		events = (result[13])
		shifts = (result[8])
		vols = (result[9])
		ntls = (result[11])
		ctms = (result[12])
		ctv = (result[14])
		#tpln_msg_txt 1-3 are old. Currently we use email to send email text and mms to send the text. tpln_msg_txt 1-3 are missing door_att
		tpln_msg_txt_1 = ('\n(1/2)\n' + str(datetime.datetime.now().strftime('%m')) + '/' + str(datetime.datetime.now().strftime('%d')) + '/' + str(datetime.datetime.now().strftime('%Y')) 
							+ '\nProgram Toplines:' + '\nDials: ' + str("{:,}".format(dials)) + '\nPhone Convos: ' + str("{:,}".format(convos)) + '\nWrong #s: ' + str("{:,}".format(w_num)) 
							+ '\nCanv Attempts: ' + str("{:,}".format(bw_att)) + '\nCanv Convos: ' + str("{:,}".format(bw_convos)) + '\n')
							
		tpln_msg_txt_2 = ('\n(2/2)\nVDRs: ' + str("{:,}".format(vdrs)) + '\nVR Forms: ' + str("{:,}".format(vr)) + '\nEvents: '+ str("{:,}".format(events)) + '\nVols: ' 
							+ str("{:,}".format(vols)) + '\nShifts: ' + str("{:,}".format(shifts)) + '\nNTLs: ' + str("{:,}".format(ntls)) 
							+ '\nCTMs: ' + str("{:,}".format(ctms)) + '\n')
		
		tpln_msg_txt_3 = ('VDRs: ' + str("{:,}".format(vdrs)) + '\nVR Forms: ' + str("{:,}".format(vr)) + '\nEvents: '+ str("{:,}".format(events)) + '\nVols: ' 
							+ str("{:,}".format(vols)) + '\nShifts: ' + str("{:,}".format(shifts)) + '\nNTLs: ' + str("{:,}".format(ntls)) 
							+ '\nCTMs: ' + str("{:,}".format(ctms)) + '\n')		
							
		tpln_msg_email = ('\nGood morning, today\'s PTR is attached.'
							+ '\n\n\tDials: ' + str("{:,}".format(dials)) + '\n\tPhone Conversations: ' + str("{:,}".format(convos)) + '\n\tWrong Numbers: ' + str("{:,}".format(w_num)) 
							+ '\n\tCanvass Attempts: ' + str("{:,}".format(bw_att)) + '\n\tCanvass Conversations: ' + str("{:,}".format(bw_convos))  
							+ '\n\tUnique Door Attempts: ' + str("{:,}".format(door_att)) + '\n\tVDRs: ' + str("{:,}".format(vdrs)) + '\n\tVR Forms: ' + str("{:,}".format(vr))
							+ '\n\tCommit to Vote Cards: ' + str("{:,}".format(ctv)) + '\n\tEvents: '+ str("{:,}".format(events)) + '\n\tVolunteers: ' 
							+ str("{:,}".format(vols)) + '\n\tShifts: ' + str("{:,}".format(shifts)) + '\n\tNTLs: ' + str("{:,}".format(ntls)) 
							+ '\n\tCTMs: ' + str("{:,}".format(ctms))
							+ '\n\nThanks,\nBill' + '\n\nGet the statewide toplines texted straight to your cellphone every morning. Sign up here: http://bit.ly/BGTXTextSignup')
							
		tpln_msg_mms = ('\n' + str(datetime.datetime.now().strftime('%m')) + '/' + str(datetime.datetime.now().strftime('%d')) + '/' + str(datetime.datetime.now().strftime('%Y')) 
							+ '\nProgram Toplines:' + '\nDials: ' + str("{:,}".format(dials)) + '\nPhone Convos: ' + str("{:,}".format(convos)) + '\nWrong #s: ' + str("{:,}".format(w_num)) 
							+ '\nCanv Attempts: ' + str("{:,}".format(bw_att)) + '\nCanv Convos: ' + str("{:,}".format(bw_convos)) + '\nUnique Doors: ' + str("{:,}".format(door_att)) + '\nVDRs: ' 
							+ str("{:,}".format(vdrs)) + '\nVR Forms: ' + str("{:,}".format(vr)) + '\nCTV Cards: ' + str("{:,}".format(ctv)) + '\nEvents: '+ str("{:,}".format(events)) + '\nVols: ' 
							+ str("{:,}".format(vols)) + '\nShifts: ' + str("{:,}".format(shifts)) + '\nNTLs: ' + str("{:,}".format(ntls)) 
							+ '\nCTMs: ' + str("{:,}".format(ctms)) + '\n')
							
	else:
		pass
			
w.close()

print str(datetime.datetime.now().strftime('%Y%m%d %H:%M:%S')) + ':\tprogram toplines report dataset complete'

# build excel report
root = os.path.join('C:/Users/bgtx.bward/Desktop/BGTX/Reports/Program History Report/')
file_name = 'Program Toplines Report _TEMPLATE_MACRO_1.xlsm'
file = os.path.join(root,file_name)

print str(datetime.datetime.now().strftime('%Y%m%d %H:%M:%S')) + ':\tgenerating excel report'

print str(datetime.datetime.now().strftime('%Y%m%d %H:%M:%S')) + ':\topening excel'

myExcel = Dispatch('Excel.Application')
myExcel.Workbooks.Close()

# no screen flashes -- redundancy, turned off in VBA code
myExcel.Visible = 0

print str(datetime.datetime.now().strftime('%Y%m%d %H:%M:%S')) + ':\topening workbook'
myExcel.Workbooks.Open(file)

# must wrap file name in single quotes if file name contains spaces
print str(datetime.datetime.now().strftime('%Y%m%d %H:%M:%S')) + ':\trunning macro'
myExcel.Run("'" + file + "'" +'!close_and_save')

# no pop up prompts when deleting tabs or saving as. Redundancy, they are turned off in the VBA code. If doing in VBA, just remember to switch it back to 'true' in vba as last line of macro.
myExcel.DisplayAlerts = 0
myExcel.Workbooks.Close()

"""
	Program Report SQL Run and File Creation Module End
"""



""" 
	Text Subscription Module Begin
"""

#RETRIEVE SUBSCRIBER GOOGLE DOC
#set retrieval credentials
username = "google_user"
password = "google_pw"
#Name of Doc in my Drive: "Program Report Text Subscription Signup (Responses)"
docid = "google_doc_id"

client = gspread.login(username, password)
spreadsheet = client.open_by_key(docid)

#saves each worksheet in the google doc as a separate csv. We care about i=0, the first tab of the sheet, hence the if statement.
for i, worksheet in enumerate(spreadsheet.worksheets()):
	if i==0:
		#if you just run this, you'll get all the worksheets in the google doc
		filename = 'C:/Users/bgtx.bward/Desktop/BGTX/python_scripts/reports/program_report/automation/csv_save/' + docid + '-worksheet' + str(i) + '.csv'
		with open(filename, 'wb') as f:
			writer = csv.writer(f)
			writer.writerows(worksheet.get_all_values())
	else:
		pass

#CREATE PYTHON DICTIONARY
#setting csv_file to the csv we just made when we downloaded the google doc
csv_file = csv.reader(open('filename_from_above_step'))

#create dictionary
subscribers = {}

#firstline lets you run an if statement in your for loop to ignore the header. I'm sure there are more elegant ways, but this was first code I found on the interwebs. It works.
firstline = True

#load each row of the csv into the dictionary
for row in csv_file:
	if firstline:
		firstline = False
		continue
	#use the first column as the key	
	key = row[0]
	#use the rest of the rows as the values
	values = row[1:]
	#append the email addresses to the lists associated with each key. These are @mail suffixes for each carrier (looked 'em up online).
	if values[5] == 'Verizon':
		values.append(values[4] + '@vzwpix.com')
	elif values[5] == 'AT&T':
		values.append(values[4] + '@mms.att.net')
	elif values[5] == 'Sprint':
		values.append(values[4] + '@pm.sprint.com')
	elif values[5] == 'T-Mobile':
		values.append(values[4] + '@tmomail.net')
	else:
		values.append('carrier_not_supported')
	#assign values list to be the 'value' in the 'key:value' pair.
	subscribers[key] = values	

#Check your full subscriber dicitonary
#print 'Full Subscriber List\n'	
#print subscribers	

""" 
	Text Subscription Module End
"""


"""
	Topline Email w Attachment Send Module Begin
"""

# Sending toplines completion email to listserv

#define variables for email send
toaddrs = 'recipient_listserv'
root_email = os.path.join('C:/Users/bgtx.bward/Desktop/BGTX/Reports/Program History Report/scripted_outputs/')
timestamp_email = datetime.datetime.now().strftime('%Y') + '.' + str(int(datetime.datetime.now().strftime('%m'))) + '.' + str(int(datetime.datetime.now().strftime('%d')))
timestamp_email_subject = datetime.datetime.now().strftime('%m') + '/' + datetime.datetime.now().strftime('%d') + '/' + datetime.datetime.now().strftime('%Y')
file_name_email = 'C:/Users/bgtx.bward/Desktop/BGTX/Reports/Program History Report/Report/' + timestamp_email + '_Program Toplines Report.xlsm'
file_email = os.path.join(root,file_name)
subject = 'Program Toplines Report - ' + timestamp_email_subject
file  = 'C:/Users/bgtx.bward/Desktop/BGTX/Reports/Program History Report/Report/2014.8.19_Program Toplines Report.xlsm'
gmail_user = "gmail_user"
gmail_pwd = "gmail_pw"

#define the function that will send the email
def mail(to, subject, text, attach):
   msg = MIMEMultipart()

   msg['From'] = gmail_user
   msg['To'] = to
   msg['Subject'] = subject

   msg.attach(MIMEText(text))

   part = MIMEBase('application', 'octet-stream')
   part.set_payload(open(attach, 'rb').read())
   Encoders.encode_base64(part)
   part.add_header('Content-Disposition',
           'attachment; filename="%s"' % os.path.basename(attach))
   msg.attach(part)

   mailServer = smtplib.SMTP("smtp.gmail.com", 587)
   mailServer.ehlo()
   mailServer.starttls()
   mailServer.ehlo()
   mailServer.login(gmail_user, gmail_pwd)
   mailServer.sendmail(gmail_user, to, msg.as_string())
   # Should be mailServer.quit(), but that crashes...
   mailServer.close()

  

#send the email  
mail(toaddrs, subject, tpln_msg_email, file_name_email)

print 'sent the email to "recipient_listserv"'

"""
	Topline Email w Attachment Send Module End
"""


""" 
	Text Send Module Begin
"""

#SEND PROGRAM TOPLINES VIA TEXT TO VALID TEXT SUBSCRIBERS
# Remove: keys without unique ids, people with program report subscription: "No", people without approved email address, people not on Verizon, AT&T, Sprint or T-Mobile, etc.
print '\nTexting Valid Subscribers:'	
for key in subscribers:	
	if subscribers[key][6] == 'No':
		pass
	elif key == '':
		pass
	elif subscribers[key][9] == 'N':
		pass
	elif subscribers[key][10] == 'carrier_not_supported':
		pass
	else:
		username = 'sender_email'
		password = 'sender_pw'
		fromaddr = 'sender_email'
		toaddrs  = subscribers[key][10]
		#msg = '[USE THIS TO SUB AN ALTERNATE MSG TO THIS GROUP]'
	
		server = smtplib.SMTP('smtp.gmail.com:587')
		server.starttls()
		server.login(username,password)
		#server.sendmail (fromaddr, toaddrs, msg) #for alternate message send
		server.sendmail(fromaddr, toaddrs, tpln_msg_mms)
		server.quit()
		
		print str(datetime.datetime.now().strftime('%Y%m%d %H:%M:%S')) + '\ttexted ' + subscribers[key][1] + ' ' + subscribers[key][2] + ' at ' + subscribers[key][10]

""" 
	Text Send Module End
"""


