#!/usr/bin/env python

import imaplib
import email
import sys
import datetime
import smtplib
from email.mime.text import MIMEText
import types


def getFirstTextPart(msg):
    maintype = msg.get_content_maintype()
    if maintype == 'multipart':
        for part in msg.get_payload():
            if part.get_content_maintype() == 'text':
                return part.get_payload()
    elif maintype == 'text':
        return msg.get_payload()


def fetchmailbox(account, password, numberOfEmails):
	print 'Fetching mailbox: ' + str(account)
	mail = imaplib.IMAP4_SSL('imap.emailsrvr.com')
	mail.login(account, password)
	mail.list()
	status, count = mail.select('inbox')
	messages = str(count)
	messages = messages.replace('[', '')
	messages = messages.replace(']', '')
	messages = messages.replace('\'', '')
	print messages + ' messages in inbox'
	varMailBoxText = '<p>There are ' + str(messages) + ' messages in this mailbox.</p>'

	typ, data = mail.search(None, 'ALL')
	ids = data[0]
	id_list = ids.split()
	#get the most recent email id
	try:
		latest_email_id = int (id_list[-1])
	except IndexError:  
		numberOfEmails = 0
	print 'id_list ' + str(id_list)
	#iterate through 'numberOfEmails' messages in decending order starting with latest_email_id
	#the '-1' dictates reverse looping order
	if numberOfEmails > 0:
		try:
			for i in range(latest_email_id, latest_email_id - numberOfEmails, -1 ):
				try:
					print 'Fetching email ' + str(i)
					typ, data = mail.fetch(i, '(RFC822)')
				except:
					print 'Error while fetching emails.'
					
				for response_part in data:
					if isinstance(response_part, tuple):
						msg = email.message_from_string(response_part[1])
						varSubject = msg['subject']
						varFrom = msg['from']
						varDate = msg['date']
						varBody = getFirstTextPart(msg)
					#Remove the brackets around the sender email address
					varFrom = varFrom.replace('<', '')
					varFrom = varFrom.replace('>', '')

					#add ellipsis (...) if subject length is great than 35 characters
					if len(varSubject) > 35:
						varSubject = varSubject[0:32] + '...'
					if isinstance(varBody, str):
						if len(varBody) > 200:
							varBody = varBody[0:200] + '...'
						if varBody[0:15] == 'BEGIN:VCALENDAR':
							varBody = 'Message is a calendar item.'
					if isinstance(varBody, types.NoneType):
						varBody = 'Message cannot currently be read.'
				# varDate format is Thu, 2 Jan 2014 15:46:54 	
				# Add leading 0 to day if single digit
				if varDate[6:7] == ' ':
					varDate = varDate[:5] + '0' + varDate[5:] # Insert leading 0 at 5th character of string
				# Remove Time zone text if present (e.g. '(EDT)')
				if varDate[-1] == ')':
					varDate = varDate[:-6]
				#Remove time zone and ' ' char (e.g. ' -0000')
				varDate = varDate[:-6]
				try	:
					emailDateTime = datetime.datetime.strptime(varDate, "%a, %d %b %Y %H:%M:%S")
				except ValueError:
					print 'Value Error in date conversion'

				emailDate = emailDateTime.date() #Create date type
				print 'Email Date = ' + str(emailDate)
				print 'Current Date = ' + str(currentDate)
				emailAge = abs((currentDate - emailDate).days)
				print 'Message is ' + str(emailAge) + ' days old'
				strDays = 'days'
				if emailAge == 1:
					strDays = 'day'
				
				strPara = '<p>'
				if emailAge < highlightAge:
					strPara = '<p style="background-color : yellow;">'
				varEmail = strPara + varDate + ' - '+ str(emailAge) + ' ' + strDays + ' old. ' + ' [' + varFrom.split()[-1] + '] ' + varSubject + '</p>'
				if emailAge < highlightAge:
					varEmail = varEmail + '<p style="font-size: small">' + str(varBody) + '</p>'
				varMailBoxText = varMailBoxText + '\n' + varEmail
					
		except IndexError:  
			print 'Index error while fetching emails.'
	
	mail.logout
	print varMailBoxText
	return varMailBoxText;

	

# Global Variables
highlightAge = 7
systemEmailAddress = 'noreply@example.com'
receiverEmailAddress = 'user@example.com'
	
# Get current date
currentDate = (datetime.date.today())

# Init messageBody for HTML
messageBody = '<body style="font-family : Arial, Helvetica, sans-serif">' + '<p>Here is your daily digest of the recent emails in-boxes. Messages sent in the last ' + str(highlightAge) + ' days are highlighted.</p>'


users = {} # Create a mapping for the list of users
users["email1@example.com"] = "password"
users["email2@example.com"] = "password"


for user, pw in users.iteritems():
	messageBody = messageBody + '<div class="user"><h3>' + user + '</h3>'
	messageBody = messageBody + fetchmailbox(user, pw, 3) + '</div>'

messageBody = messageBody + '</body>'	
	
print messageBody

msg = MIMEText(messageBody, 'html')
msg['Subject'] = 'Email Digest'
msg['From'] = systemEmailAddress
msg['To'] = receiverEmailAddress

s = smtplib.SMTP('localhost')
s.sendmail(systemEmailAddress, [receiverEmailAddress], msg.as_string())
s.quit()
