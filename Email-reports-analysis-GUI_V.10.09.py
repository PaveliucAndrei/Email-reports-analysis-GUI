# Import the relevant libraries
from win32com.client import Dispatch, GetActiveObject
from win32com.shell import shell, shellcon # type: ignore
from pathlib import Path
from datetime import date
from shutil import make_archive
from json import load
import os
import re
# GUI librarie imports 
import customtkinter as ct
from CTkTable import *

# Emails credentials
with open('creds.json') as creds:
    CREDENT = load(creds)

WMSERVICE_EMAIL = CREDENT['WMSERVICE_EMAIL']
XAASIT_EMAIL = CREDENT['XAASIT_EMAIL']

# Import Clients config file
with open('clients_config.json') as clients_config:
    CLIENTS = load(clients_config)

XAASIT_FOLDER = 'xaas-it (PDM)'

# Email sub folders
CHECK_ERROR = '1_Check_Error'

# Create output folder
# USSER_DESKTOP_PATH = os.path.join(os.environ['USERPROFILE'], 'Desktop')
USSER_DESKTOP_PATH = shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOP, None, 0)
WK_DIR = Path(USSER_DESKTOP_PATH) / 'XaasIT_WK_DIR'
WK_DIR.mkdir(parents=True, exist_ok=True)
DATE = date.today()
DATE = DATE.isoformat()

ARCHIVE_THRESHOLD = 25

# Subject patterns
SWIFT_PATTERN = r'SWIFT+.*Delivery: (.*_[0-9]{4}-[0-9]{2}-[0-9]{2}-[0-9]{2}:[0-9]{2}:[0-9]{2}) on Server:'
USUAL_DELIVERYS_PATTERN = r'Delivery: (.+) on Server:'

# @@@ For testing @@@
USER_EMAIL = CREDENT['USER_EMAIL']
EMAILS_LIMIT = 500
TEST_XAASIT = 'Test Xaas'
CHECK_OK = '2_Check_OK'

#--------------------------------------------------------------------------------------

def email_connection(user_email_address:str, sub_folder:str, main_folder = 'Inbox'):
    # Connect to email
    OUTLOOK = get_running_outlook_inst()
    if not get_running_outlook_inst():
        OUTLOOK = Dispatch('Outlook.Application')
    
    for i in range(10): # Try to connect to OUTLOOK NameSpace. Thie loop is needed because of the: "AttributeError - Outlook.Application.GetNameSpace" thet some times appears
        try:            
            OUTLOOK_NAMESPACE = OUTLOOK.GetNameSpace('MAPI')
        except AttributeError as Outlook_NameSpace_err:
            print(f'### Erron on run: {i} -- {Outlook_NameSpace_err}')
        else:
            break
    # Connect to sub email folders
    FOLDER = OUTLOOK_NAMESPACE.Folders(user_email_address).Folders(main_folder).Folders(sub_folder)    
    # Get the emails
    emails = FOLDER.Items
    # Sort messages by ReceivedTime (descending order for most recent)
    emails.Sort('[ReceivedTime]', 1) # xlAscending = 1 # xlDescending = 2
    
    return OUTLOOK, emails

def subject_sn_filter(emails, sn_mask:str) -> list:

    email_obj = []

    i = 0
    for email in emails:
        subject = email.Subject

        if sn_mask in subject:
            email_obj.append(email)

        if i >= EMAILS_LIMIT:
            break
        i += 1

    return email_obj

def subject_mask_filter(subject:str) -> str | str:

    swift_match = re.search(SWIFT_PATTERN, subject)
    usnual_deliverys_match = re.search(USUAL_DELIVERYS_PATTERN, subject)

    if swift_match is not None:
        return 'SWIFT', swift_match.group(1)
    elif usnual_deliverys_match is not None:
        return 'usually_delivery', usnual_deliverys_match.group(1)

def extract_attachments(attachments, target_folder:Path) -> int | list:

    attachments_count = 0
    attachments_extracted = []
    # Iterate to each attachments
    for attachmet in attachments:
        # Download the attachments from the email
        attachmet.SaveAsFile(target_folder / str(attachmet))
        attachments_count += 1
        # Store the name of the files
        attachments_extracted.append(attachmet.FileName)

    return attachments_count, attachments_extracted
    
# Left join list
def left_join(swifts_name_subject, attachments_extracted) -> list:
    # Convert the data to sets
    swifts_name_subject = set(swifts_name_subject)
    attachments_extracted = set(attachments_extracted)

    return list(swifts_name_subject - attachments_extracted)

def make_email_body(swifts_list:list, regular_deliveries:list) -> str:

    swifts_list.sort()
    regular_deliveries.sort()
    
    introduction = 'Good day,\n\n'
    swift_deliveries = 'The next SWIFTs are processed with errors:\n\n'
    normal_deliveries = 'The following interface deliveries were processed with errors:\n\n'
    final = 'Please check.\n\nBest regards.\n\n\n'

    body = introduction
    # SWIFTs block
    if swifts_list:
        body += swift_deliveries
        for swift in swifts_list:
            body += f'{swift}\t\t\n'
        body += '\n'
    # Regular deliveries block
    if regular_deliveries:
        body += normal_deliveries
        for delivey in regular_deliveries:
            body += f'{delivey}\t\t\n\n\n'

    body += final

    return body

def make_email(outlook, em_from, em_to, em_bbc, em_subject, swifts_list, usually_deliveries, attachments_path:Path):
    # Create a new email message
    email = outlook.CreateItem(0)  # 0 = create email item
    
    # Set the email properties
    if isinstance(em_to, list): 
        email.To = ';'.join(em_to)
    else: 
        email.To = em_to
    email.BCC = em_bbc
    email.Subject = em_subject
    # Set the email body
    email.BodyFormat = 1 # 2 for olFormatHTML https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa219371(v=office.11)?redirectedfrom=MSDN
    email.Body = make_email_body(swifts_list, usually_deliveries)
    # Set the sender's email account
    email.SentOnBehalfOfName = em_from
    
    # Add attachments
    if attachments_path.suffix != '.zip':
        for item in os.listdir(attachments_path):
            # Check if current file_path is a file
            item_path = os.path.join(attachments_path, item)
            if os.path.isfile(item_path):
                # Attach the file
                email.Attachments.Add(item_path)
    elif attachments_path.suffix == '.zip':
        email.Attachments.Add(str(attachments_path))

    email.display()

def multiple_runs(wk_dir, client_id, date):
    extraction_counter = 1
    target_folder = wk_dir / client_id / date
    if target_folder.exists():
        # Pars the directories name, as a list
        dir_list = [p for p in Path(target_folder).parent.iterdir() \
                    if p.is_dir() and '_' in p.name and p.name.split('_')[0] == date]
        # Extract the number of runs, as a list
        extraction_counter_list = [run_count for d in dir_list if (run_count := int(d.name.split('_')[1])) >= 1]
        
        if extraction_counter_list:
            extraction_counter = max(extraction_counter_list) + 1
            target_folder = target_folder.with_name(target_folder.name + f'_{extraction_counter}')
        else:
            target_folder = target_folder.with_name(target_folder.name + f'_{extraction_counter}')

    target_folder.mkdir(parents=True, exist_ok=True)
    
    return target_folder

def open_folder(path, missing_folder_message):

    if not Path.exists(path):  # Check if folder path is valid
        missing_folder_message.configure(text=f'Error: Folder path "{path}" does not exist.')
        # raise Exception(f'Error: Folder path "{path}" does not exist.')
        return None

    return os.startfile(path)

def get_running_outlook_inst():
    try:
        return GetActiveObject('Outlook.Application')
    except Exception:
        return False

def open_outlook():
    os.startfile('outlook')

#--------------------------------------------------------------------------------------

def main():
    
    # +++ GUI
    # Set the theme
    ct.set_appearance_mode('System')  # Modes: system (default), light, dark
    ct.set_default_color_theme('dark-blue')  # Themes: blue (default), dark-blue, green

    # Root
    ROOT= ct.CTk()
    # Window configuration
    ROOT.title('Attachments Extraction')
    ROOT.geometry('650x400')

    ROOT.grid_columnconfigure(0, weight=1)
    ROOT.grid_rowconfigure(0, weight=1)
    
    # Frames
    FRAME_EXTRACT = ct.CTkFrame(ROOT, width=550, height=360, border_width=2)
    FRAME_EXTRACT.grid(row=0, column=0, padx=(10, 10), pady=10, sticky='NSWE')
    
    radio_selection = ct.StringVar(value='')

    def extracting():
        # Some variables
        attachment_count = 0
        deliverys = {'SWIFTs_list':[],
                    'usually_deliveries':[]}
        attachments_extracted = []
       
        # Filter for the slected client
        check_error_client = CLIENTS[radio_selection.get()]
        if not check_error_client:
            return MESSAGE_SELECT.configure(text='Please make a selection')
        
        # Extract the SN email, for the selected client
        sn_emails = subject_sn_filter(xaasit_emails, check_error_client['SN_mask'])

        # 📂 Create separate folder for each client\day and for multiple runs
        target_folder = multiple_runs(WK_DIR, check_error_client['client_ID'], DATE)

        # Filter for the name of the SWIFTs form the email subject and extract the attachments
        for sn_email in sn_emails:        
            tip, delivery_no = subject_mask_filter(sn_email.Subject)
            # Separete the SWIFT delivereis and normal ones
            if delivery_no is not None:
                if tip == 'SWIFT':
                    deliverys['SWIFTs_list'].append(delivery_no)
                elif tip == 'usually_delivery':
                    deliverys['usually_deliveries'].append(delivery_no)
            # Extract the attachments
            att_count, att_extracted = extract_attachments(sn_email.Attachments, target_folder)
            attachment_count += att_count
            attachments_extracted.extend(att_extracted)

            # Get the missing SWIFTs. The ones that are not attached to the emails
            # SWIFT_unextract = left_join(deliverys['SWIFTs_list'], attachments_extracted)

        # Archive the attachments
        if attachment_count > ARCHIVE_THRESHOLD:
            archive_path = make_archive(target_folder, 'zip', target_folder)
            # archive_path = make_archive(target_folder / target_folder.stem, 'zip', target_folder.parent, target_folder.name)
            target_folder = Path(archive_path)

        MESSAGE_SELECT_TEXT = f'{str(check_error_client['client_ID'])}: {attachment_count} attachments extracted'
        MESSAGE_FILE_LOC_TEXT = f'File location -> {target_folder}'

        MESSAGE_SELECT.configure(text=MESSAGE_SELECT_TEXT)
        MESSAGE_FILE_LOC.configure(text=MESSAGE_FILE_LOC_TEXT)
        BUTTON_GO_TO = ct.CTkButton(FRAME_EXTRACT, 
                                    text='Go to 📂', 
                                    command=lambda: open_folder(target_folder, MESSAGE_FILE_LOC), 
                                    font=('Helvetica', 22))
        BUTTON_GO_TO.place(relx=0.3, rely=0.9, anchor=ct.N)
        BUTTON_EMAIL = ct.CTkButton(FRAME_EXTRACT,
                                    text='📧',
                                    command=lambda: make_email(OUTLOOK,
                                                                XAASIT_EMAIL,
                                                                check_error_client['eMail_To'],
                                                                XAASIT_EMAIL,
                                                                check_error_client['client_ID'] + ' - ' + check_error_client['eMail_subject'],
                                                                deliverys['SWIFTs_list'],
                                                                deliverys['usually_deliveries'],
                                                                target_folder), 
                                    font=('Helvetica', 22))
        BUTTON_EMAIL.place(relx=0.7, rely=0.9, anchor=ct.N)

    # Create the main UI elements
    MESSAGE_OPENING = ct.CTkLabel(FRAME_EXTRACT, text='', font=('Halvetica', 27))
    MESSAGE_OPENING.place(relx=0.5, rely=0.1, anchor=ct.N)
    BUTTON_MAIN = ct.CTkButton(FRAME_EXTRACT, font=('Helvetica', 22))

    # Connect to email and extrat the emails list
    # OUTLOOK, test_xaasit_emails = email_connection(USER_EMAIL, TEST_XAASIT)
    OUTLOOK, xaasit_emails = email_connection(XAASIT_FOLDER, CHECK_ERROR)

    # Opening message
    MESSAGE_OPENING.configure(text='Select a client')

    # Create the radio buttons
    RADIO_AI = ct.CTkRadioButton(FRAME_EXTRACT, text='AI', value='AI', variable=radio_selection, font=('Helvetica', 22))
    RADIO_AI.place(relx=0.18, rely=0.3, anchor=ct.N)
    RADIO_EB = ct.CTkRadioButton(FRAME_EXTRACT, text='EB', value='EB', variable=radio_selection, font=('Helvetica', 22))
    RADIO_EB.place(relx=0.35, rely=0.3, anchor=ct.N)
    RADIO_KSKK = ct.CTkRadioButton(FRAME_EXTRACT, text='KSKK', value='KSKK', variable=radio_selection, font=('Helvetica', 22))
    RADIO_KSKK.place(relx=0.55, rely=0.3, anchor=ct.N)
    RADIO_SKB = ct.CTkRadioButton(FRAME_EXTRACT, text='SKB', value='SKB', variable=radio_selection, font=('Helvetica', 22))
    RADIO_SKB.place(relx=0.80, rely=0.3, anchor=ct.N)
    # Button extract
    BUTTON_MAIN.configure(text='Extract 📎', command=extracting)
    BUTTON_MAIN.place(relx=0.5, rely=0.5, anchor=ct.N)

    MESSAGE_SELECT = ct.CTkLabel(FRAME_EXTRACT, text='', font=('Halvetica', 24))
    MESSAGE_SELECT.place(relx=0.5, rely=0.7, anchor=ct.N)
    MESSAGE_FILE_LOC = ct.CTkLabel(FRAME_EXTRACT, text='', font=('Halvetica', 16))
    MESSAGE_FILE_LOC.place(relx=0.5, rely=0.8, anchor=ct.N)

    ROOT.mainloop()

if __name__ == '__main__':
    main()








