"""
Author: MAJ Crosser @ 5019

Dependencies
    - pywin32
        - from powershell (may require admin powershell)
        - pip install pywin32
    - Text file containing the message for the body of the email in same dir
        - "email_body_message.txt"
    - Text file containing the subject of the email in same dir
        - "email_subject.txt"
    - CSV file linking files to email addresses. Reads file name up to the underscore for dictionary key
        - "roster.csv"
        
Usage
    - change directory below to path of files you wish to email
    - place an "email_body_message.txt", "email_subject.txt", and "roster.csv" files in the director
    - execute script

Output
    - emails files to cadets (prints msg for each)
    - Moves any emailed files to a folder named "emailed"    
"""

from win32com.client import Dispatch, constants
from os.path import abspath, isdir
from os import listdir, makedirs, rename, remove, getcwd, walk
import csv
import sys


# Change the directory variable
DIRECTORY = 'C:\\Users\\david.niblick\\PycharmProjects\\mailer\\output\\'

# Append any files you wish to except from the mailing list
EXCEPTION_LIST = []
EXCEPTION_LIST.append('grades.xlsx')

EMAIL_BODY_FILE_NAME = DIRECTORY + 'email_body_message.txt'
CLASS_EMAIL_ROSTER_FILE_NAME = DIRECTORY + 'roster.csv'
EMAIL_SUBJECT_FILE_NAME = DIRECTORY + 'email_subject.txt'
EXCEPTION_LIST.append('email_body_message.txt')
EXCEPTION_LIST.append('roster.csv')
EXCEPTION_LIST.append('email_subject.txt')


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
            key = row[0]
            email_dict[key] = row[1]
    return email_dict


print(CLASS_EMAIL_ROSTER_FILE_NAME)
ed = get_email_dict()

print(ed)
total_files = 0
file_count = 0


makedirs(DIRECTORY + '/' + 'emailed', exist_ok=True)

with open(EMAIL_BODY_FILE_NAME) as body_handle:
    body_msg = body_handle.read()

with open(EMAIL_SUBJECT_FILE_NAME) as body_handle:
    email_subject = body_handle.read()

if body_msg == '' or email_subject == '':
    print('Missing valid email body from email_body_message.txt or subject from email_subject.txt')
    exit()

files = [item for item in listdir(DIRECTORY) if item.endswith('.txt')]
total_files += len(files) - 2

for file_name in files:
    if file_name not in EXCEPTION_LIST:
        print("file_name is ", file_name)
        username = file_name.split('_')[0]
        print(username)
        print(f"Emailing {file_name} to {ed[username]}")
        send_email(ed[username],
                   body_msg,
                   email_subject,
                   DIRECTORY + '/' + file_name)
        file_count += 1
        try:
            rename(DIRECTORY + '/' + file_name,
                   DIRECTORY + '/emailed/' + file_name)
        except FileExistsError:
            remove(DIRECTORY + '/' + file_name)

print("total files: ", total_files)
print("total emails: ", file_count)
