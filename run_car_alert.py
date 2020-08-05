from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import time
from string import Template
import smtplib
import yaml
import ssl
import datetime


SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SPREADSHEET_ID_KEY = 'spreadsheet_id'
RANGE_NAME_KEY = 'range_name'


def get_config(path_to_file: str) -> dict:
    with open(path_to_file) as fid:
        return yaml.load(fid)


def time_in_range(start, end, x):
    """Return true if x is in the range [start, end]"""
    if start <= end:
        return start <= x <= end
    else:
        return start <= x or x <= end

def main():

    config_dict = get_config('run_car_alert_config.yaml')
    start_time = datetime.time(6, 0, 0)
    end_time = datetime.time(23, 0, 0)
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

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    old_values = []
    sheet = service.spreadsheets()
    while True:
        now_time = datetime.datetime.now().time()
        if time_in_range(start=start_time, end=end_time, x=now_time):
            result = sheet.values().get(spreadsheetId=config_dict.get(SPREADSHEET_ID_KEY),
                                        range=config_dict.get(RANGE_NAME_KEY)).execute()

            values = result.get('values', [])
            if not values:
                print('No data found.')
            else:
               old_values = my_bl(values=values, old_values=old_values, config_dict=config_dict)
            print('{} - in time range'.format(now_time))
        else:
            print('{} - out of time range'.format(now_time))
        time.sleep(60)



def my_bl(values: list, old_values:list, config_dict: dict) -> list:
    if values[0] != old_values:
        print('old file {}'.format(old_values))
        print('new file {}'.format(values[0]))
        old_values = values[0]
        send_mail(config_dict=config_dict)

    return  old_values


def get_contacts(filename):
    names = []
    emails = []
    with open(filename, mode='r', encoding='utf-8') as contacts_file:
        for a_contact in contacts_file:
            names.append(a_contact.split()[0])
            emails.append(a_contact.split()[1])
    return names, emails


def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)


def send_mail(config_dict: dict):
    gmail_user = config_dict.get('gmail_user')
    gmail_password = config_dict.get('gmail_pass')
    sent_from = gmail_user
    to = config_dict.get('mail_list')
    subject = 'new car spa dates'
    body = 'new car spa dates'
    email_text = """\
    From: {}
    To: {}
    Subject: {}

    {}
    """.format (sent_from, ", ".join(to), subject, body)
    context = ssl.create_default_context()
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.ehlo()
        server.starttls(context=context)
        server.ehlo()
        server.login(gmail_user, gmail_password)
        server.sendmail(sent_from, to, email_text)
        server.close()



if __name__ == '__main__':
    main()