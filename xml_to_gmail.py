#! python3
from __future__ import print_function
import pickle
import os.path
import base64
import sys
import time
from xlrd import open_workbook
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import mimetypes

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/gmail.send']

def parse_xls(service, filename):
    wb = open_workbook(filename)
    mail_from = "me"
    header_row = 2
    for s in wb.sheets():
        print ('Sheet:',s.name)
	# skip Sheet
        if s.name != "electrical":
            continue
        header = []
        #entire_sheet_out = ""
        for row in range(s.nrows):
            if row < header_row:
                continue
            values = []
            for col in range(s.ncols):
                values.append(s.cell(row,col).value)
    	    # read header
            if row == header_row:
                header = values
                continue

            name = values[1]
            email_to = values[2]
            # skip rows with no email
            if email_to == "":
                print ("skipping empty line")
                continue
            # format key n value
            #res = [str(i) + ": " + str(j) for i, j in zip(header, values)]
            row_out  = "<p>Hi " + name + ",</p>"
            row_out += "<p>Regards,<br>NAME<br><br></p>"

            #row_out += '\n'.join(res) + "\n"
            #print (row_out)
            #entire_sheet_out += row_out
            mail_subject = "your subject"
            message = create_message(mail_from, email_to, mail_subject, row_out)
            result = send_message(service, mail_from, message)
            print (result)
            # rate limits
            time.sleep(1)
            print ("mail sent to " + email_to)
        #message = create_message(mail_from, mail_to, mail_subject, entire_sheet_out)
        #result = send_message(service, mail_from, message)
        #time.sleep(1)
        #print (result)
        #print ("mail sent" + s.name)
        #print (entire_sheet_out)

def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    filename = sys.argv[1]
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)

    # Call the Gmail API
    parse_xls(service, filename) 

def create_message(sender, to, subject, message_text):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.

  Returns:
    An object containing a base64url encoded email object.
  """
  #message = MIMEText(message_text)
  # for html encoded mail
  message = MIMEText(message_text,'html')
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject
  # does not work with python3 so its commented
  #return {'raw': base64.urlsafe_b64encode(message.as_string())}
  # debug
  #print ("mailing details: from:" + sender + " to: " + to + "sub: " + subject + " \n")
  #print ("msg: " + message_text)
  return {'raw': base64.urlsafe_b64encode(message.as_string().encode()).decode()}

def send_message(service, user_id, message):
  """Send an email message.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message: Message to be sent.

  Returns:
    Sent Message.
  """
  try:
    message = (service.users().messages().send(userId=user_id, body=message)
               .execute())
    print ('Message Id: %s' % message['id'])
    return message
  except errors.HttpError as error:
    print ('An error occurred: %s' % error)

if __name__ == '__main__':
    main()
