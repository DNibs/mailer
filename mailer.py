"""
Author: MAJ Crosser @ 5019

Dependencies
    - pywin32
        - from powershell (may require admin poweshell)
        - pip install pywin32
    - Text file containing the message for the body of the email in same dir
        - "email_body_message.txt"
    - CSV file of students names and email addresses
        - "CY305_Cadet_Roster.csv"
        
Environment
    - sync CY305 folder from SharePoint
    - navigate to "\ay202\graded_events\database_project\output"
    - run this script from your local copy
    
Usage
    from PowerShell
    py mailer.py all | [sec1] [sec2] ...
    e.g.
    > py mailer.py J25 K25
    
Output
    - searches directories and emails files to cadets (prints msg for each)
    - Moves any emailed files to a folder named "emailed"    
"""

from win32com.client import Dispatch, constants
from os.path import abspath, isdir
from os import listdir, makedirs, rename, remove, getcwd, walk
import csv
import sys

EMAIL_BODY_FILE_NAME = getcwd()+'\\admin\\email_body_message.txt'
CLASS_EMAIL_ROSTER_FILE_NAME = getcwd()+'\\admin\\roster.csv'
EMAIL_SUBJECT = "Test Mail"
USAGE = "usage: py mailer.py all | sec1 sec2 ..."
OUTPUT_DIR = "output"


def send_email(recip, body, subject, att_rel_path):
    olMailItem = 0x0
    obj = Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = subject
    newMail.Body = body
    newMail.To = recip
    attachment1 = abspath(att_rel_path)
    newMail.Attachments.Add(Source=attachment1)
    newMail.Send()


def get_email_dict():
    email_dict = {}
    
    with open(CLASS_EMAIL_ROSTER_FILE_NAME) as cSVHandle:
        email_reader = csv.reader(cSVHandle)
        for row in list(email_reader)[1:]:
            last = row[0]
            first = row[1]
            email_dict[last + '.' + first] = row[2]
    return email_dict


if len(sys.argv) < 2:
    print(USAGE)
    sys.exit()

print(CLASS_EMAIL_ROSTER_FILE_NAME)
ed = get_email_dict()
dir_list = [x[0] for x in walk(getcwd()) if OUTPUT_DIR in x[0]]

print(dir_list)
print(ed)
total_files = 0
file_count = 0
for directory in dir_list:
    if "emailed" in directory:
        continue
    print('checking directory:', directory)
    found_dir = False
    if sys.argv[1].lower() != 'all':
        for subdir in sys.argv[1:]:
            if subdir in directory:
                found_dir = True
    else:
        found_dir = True
    if not found_dir:
        continue
    print('Good directory:', directory)
    
    makedirs(directory + '/' + 'emailed', exist_ok=True)

    with open(EMAIL_BODY_FILE_NAME) as body_handle:
        body_msg = body_handle.read()

    files = [item for item in listdir(directory) if item.endswith('.txt')]
    total_files += len(files)
    
    for file_name in files:
        print("file_name is ", file_name)
        username = file_name.split('_')[0]
        print(username)
        print(f"Emailing {file_name} to {ed[username]}")
        send_email(ed[username], 
                   body_msg, 
                   EMAIL_SUBJECT, 
                   directory + '/' + file_name)
        file_count += 1
        try:
            rename(directory + '/' + file_name,
                   directory + '/emailed/' + file_name)
        except FileExistsError:
            remove(directory + '/' + file_name)

print("total files: ", total_files)
print("total emails: ", file_count)
