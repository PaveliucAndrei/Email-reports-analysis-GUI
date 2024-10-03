# Import the relevant libraries
from win32com.client import Dispatch, DispatchEx
from pathlib import Path
from datetime import date
from subprocess import Popen
import os
import re
import creds
# GUI librarie imports 
import customtkinter as ct
from CTkTable import *


USER_EMAIL = creds.USER_EMAIL
WMSERVICE_EMAIL = creds.WMSERVICE_EMAIL
XAASIT_EMAIL = creds.XAASIT_EMAIL

# Clients masks
AI = {'client':'AI',
      'SN_mask':'CHECK_ERROR AI Xentis SN',
      'eMail_To':creds.AI_eMail_To,
      'BBC':'',
      'eMail_subject':'Unprocessed deliveries'}

SKB = {'client':'SKB', 
        'SN_mask':'CHECK_ERROR SKB Xentis SN',
        'eMail_To':creds.SKB_eMail_To,
        'BBC':'',
        'eMail_subject':'Unprocessed deliveries'}

EB = {'client':'EB', 
      'SN_mask':'CHECK_ERROR EB Xentis SN',
      'eMail_To':creds.EB_eMail_To,
      'BBC':'',
      'eMail_subject':'Unprocessed deliveries'}

# Email sub folders
CHECK_ERROR = '1_Check_Error'

# Create output folder
USSER_DESKTOP_PATH = os.path.join(os.environ['USERPROFILE'], 'Desktop')
WK_DIR = Path(USSER_DESKTOP_PATH) / 'XaasIT_WK_DIR'
WK_DIR.mkdir(parents=True, exist_ok=True)
DATE = date.today()
DATE = DATE.isoformat()

# @@@ For testing @@@
EMAILS_LIMIT = 500
TEST_XAASIT = 'Test Xaas'
CHECK_OK = '2_Check_OK'

#--------------------------------------------------------------------------------------

def email_connection(user_email_address:str, sub_folder:str, main_folder = 'Inbox'):
    # Connect to email
    OUTLOOK = Dispatch('Outlook.Application')
    OUTLOOK_NameSpace = OUTLOOK.GetNameSpace('MAPI')

    # Connect to sub email folder
    FOLDER = OUTLOOK_NameSpace.Folders(user_email_address).Folders(main_folder).Folders(sub_folder)

    # Get the emails
    emails = FOLDER.Items
    # Sort messages by ReceivedTime (descending order for most recent)
    emails.Sort('[ReceivedTime]', 1) # xlAscending = 1 # xlDescending = 2

    return OUTLOOK, emails

def check_Outlook_status():
    try:   
        pass
    except AttributeError as err:
        print('Is Outlook open?')
        


def select_client(client_selected:str):

    if client_selected == 'AI': return AI
    elif client_selected == 'SKB': return SKB
    elif client_selected == 'EB': return EB

def SN_subject_email_filter(emails, client) -> list:

    email_obj = []

    i = 0
    for email in emails:

        subject = email.Subject

        if client['SN_mask'] in subject:
            email_obj.append(email)
  
        if i >= EMAILS_LIMIT: break
        i += 1

    return email_obj

def subject_mask_filter(subject):
    # Sanitize the subject
    subject = re.sub(r'[/\\:\'\'<>?|]', '', subject).strip()
    SWIFT_pattern = r'(SWIFT+.*Delivery+\s+)(.+)_([0-9]{4}-[0-9]{2}-[0-9]{2}-[0-9]{6})+'
    usnual_deliverys_pattern = r'(Delivery+\s+)(.+)_'

    SWIFT_match = re.search(SWIFT_pattern, subject)
    usnual_deliverys_match = re.search(usnual_deliverys_pattern, subject)

    if SWIFT_match is not None:
        return {'SWIFT':SWIFT_match.group(2)}
    elif usnual_deliverys_match is not None:
        return {'usually_delivery':usnual_deliverys_match.group(2)}

def extract_attachments(attachments, target_folder) -> int | list:

    attachments_count = 0
    attachments_extracted = []
    
    for  attachmet in attachments:
        # Download the attachmnts from the email
        attachmet.SaveAsFile(target_folder / str(attachmet))
        attachments_count += 1
        # Store the name of the files
        attachments_extracted.append(attachmet.FileName)

    return attachments_count, attachments_extracted
    
# Left join list
def left_join(SWIFTs_name_subject, attachments_extracted) -> list:
    # Convert the data to sets
    SWIFTs_name_subject = set(SWIFTs_name_subject)
    attachments_extracted = set(attachments_extracted)

    return list(SWIFTs_name_subject - attachments_extracted)

def make_email_body(SWIFTs_list:list, usually_deliveries:list) -> str:

    SWIFTs_list.sort()
    usually_deliveries.sort()
    
    introduction = 'Good day,\n\nThe next SWIFTs are processed with errors:\n\n'
    final = 'Please check.\n\nBest regards.\n\n\n'

    body = introduction
    for swift in SWIFTs_list:
        body += f'{swift}\t\n'
    body += '\nThe following interface deliveries were processed with errors:\n\n'
    for delivey in usually_deliveries:
        body += f'{delivey}\t\n\n\n'
    body += final

    return body

def make_email(OUTLOOK, From, To, BCC, Subject, SWIFTs_list, usually_deliveries, attachments_path):
    # Create a new email message
    email = OUTLOOK.CreateItem(0)  # 0 = create email item
    
    # Set the email properties
    if isinstance(To, list): 
        email.To = ';'.join(To)
    else: 
        email.To = To
    email.BCC = BCC
    email.Subject = Subject
    # Set the email body
    email.BodyFormat = 1
    email.Body = make_email_body(SWIFTs_list, usually_deliveries)
    # Set the sender's email account
    send_as = None
    for myEmailAddress in OUTLOOK.Session.Accounts:
        if From in str(myEmailAddress):
            send_as = myEmailAddress
            break

    if send_as != None:
        # This line basically calls the 'email.SendUsingAccount = xyz@email.com'
        email._oleobj_.Invoke(*(64209, 0, 8, 0, send_as))

    # Add attachments
    for item in os.listdir(attachments_path):
        # Check if current file_path is a file
        item_path = os.path.join(attachments_path, item)
        if os.path.isfile(item_path):
            # Attach the file
            email.Attachments.Add(item_path)

    email.display()

def open_folder(path):

    if not os.path.isdir(path):  # Check if folder path is valid
        print(f'Error: Folder path "{path}" does not exist.')

    return os.startfile(path)

#--------------------------------------------------------------------------------------

def main():
    
    OUTLOOK, Test_XaasIT_emails = email_connection(USER_EMAIL, TEST_XAASIT)
    # os.startfile('outlook')
    print(len(Test_XaasIT_emails))
    

if __name__ == '__main__':
    main()


OUTLOOK = Dispatch('Outlook.Application')
    try:
        OUTLOOK_NameSpace = OUTLOOK.GetNameSpace('MAPI')
    except AttributeError:
        print('Is Outlook open?')
    finally:
        # os.startfile('outlook')
        OUTLOOK.Visible = True

























