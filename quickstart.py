from __future__ import print_function

import os.path
from email.mime.text import MIMEText

import base64
import pandas as pd
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# defining scope to request specific permissions
SCOPES = ['https://mail.google.com/']

# Logic code to send messages
def send_message(service, user_id, message):
#service: Authorized Gmail API service instance.
#user_id: email address of Authorised user.    
  try:
    message = service.users().messages().send(userId=user_id, body=message).execute()
    print("Message sent successfully")
    print('Message Id: %s' % message['id'])
    return message
  except Exception as e:
    print('An error occurred: %s' % e)
    return None
#Logic code to create messages
def create_message(sender, to, subject, message_text):
#Sender: email address of sender
#to: email address of receiver
#subject: subject of email
#message_txt: message to send
  message = MIMEText(message_text)
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject
  raw_message = base64.urlsafe_b64encode(message.as_string().encode("utf-8"))
  return {
    'raw': raw_message.decode("utf-8")
  }
#Logic code to search for matching messages.
#Attention: this function only return message IDs of matching messages.

def MatchingContent(service, user_id, keyword=''):
#service: Authorized Gmail API service instance.
#user_id: email address of Authorised user.    
#keyword: to filter messages    
    try:
        response = service.users().messages().list(userId=user_id,
                q=keyword).execute()
        messages = []
        if 'messages' in response:
            messages.extend(response['messages'])
        
        while 'token' in response:
            page_token = response['Token']
            response = service.users().messages().list(userId=user_id,
                    q=keyword, pageToken=page_token).execute()
            messages.extend(response['messages'])
        return messages
    except Exception as e:
        print('An error occurred: %s' % e)
        return None

#logic code to get content of message from the message ID.        
def GetMessage(service, user_id, msg_id):
    try:
        #full: Returns the full email message
        full_message = service.users().messages().get(userId=user_id, id=msg_id,
                format='full').execute()
        return full_message
    except Exception as e:
        print('An error occurred: %s' % e)
        return None

#Logic code to collect all messages in one list.          
def GetListofMessages(matching_message,service, sender):
    df = pd.DataFrame()
    
    for i in matching_message:
        complete_message = GetMessage(service, sender, i['id'])
        message_dict = dict(complete_message)
        d_msg = pd.DataFrame([message_dict])
        df = pd.concat([df,d_msg], ignore_index=True)      
    return df         

def main():

    creds = None
    #token.json contains details for user's access and creates automatically for the first time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    #To allow user to login if credentials are not valid..
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0) 
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)
        #Calling function to create message
        message = create_message('kevinaiwala1998@gmail.com','kevinaiwala1998@gmail.com', "hello there", "You got 1 billion dollar bumper price" )
        #Calling function to send message
        send_message(service, "kevinaiwala1998@gmail.com", message)
        #Calling function to search for matching mesaages and to get message IDs of matching messages 
        matching_message = MatchingContent(service,
            'kevinaiwala1998@gmail.com', keyword='hello')
        #Calling fuction to get all matching messages in one list.   
        content = GetListofMessages(matching_message,service,"kevinaiwala1998@gmail.com")
        content.rename(columns={'snippet':'Message'},inplace=True)
        try:
            content2= content["Message"]
            
        except:
            pass 
        #Creating csv file with content of matching messages.  
        content2.to_csv('MatchingMessages.csv',index=False)
        print("CSV File created successfully") 
            
    except HttpError as error:
        #Error handler for Gmail API
        print(f'An error occurred: {error}')


if __name__ == '__main__':
    main()