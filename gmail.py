
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from bs4 import BeautifulSoup
import base64
import re
import pandas as pd

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def full(payload):
    try:
        failed_email = mail_delivery_1(payload)
        if failed_email is not None:
            return failed_email
    except:
        pass
    try:
        failed_email = mail_delivery_2(payload)
        if failed_email is not None:
            return failed_email
    except:
        pass
    try:
        failed_email = failure_notice(payload)
        if failed_email is not None:
            return failed_email
    except:
        pass
        # failed_email = None
        # return failed_email

def postmaster_1(payload):
    try:
        for parts_one in payload['parts']:
            if parts_one['partId'] == '2':
                for parts_two in parts_one['parts']:
                    head = parts_two['headers']
                    for a in head:
                        if a['name'] == 'BCC' or a['name'] == 'Bcc':
                            failed_email = a['value']
                            return failed_email
                        else:
                            pass
    except:
        return None


def postmaster_2(txt):
    failed_email = re.findall(r'[\w.+-]+@[\w-]+\.[\w.-]+', txt['snippet'])
    return failed_email[0]


def mail_delivery_1(payload):
    try:
        for i in payload['headers']:
            if i['name'] == 'X-Failed-Recipients':
                failed_email = i['value']
                return failed_email
        failed_email = postmaster_1(payload)
        return failed_email
    except:
        return None


def mail_delivery_2(payload):
    try:
        for parts_one in payload['parts']:
            if parts_one['partId'] == '2':
                for parts_two in parts_one['parts']:
                    head = parts_two['headers']
                    for a in head:
                        if a['name'] == 'Received':
                            line = a['value']
                            emails = re.findall(r'[\w.+-]+@[\w-]+\.[\w.-]+', line)
                            return emails[0]
    except:
        return None
        # pass


def text(payload):
    data = None
    parts_one = payload.get('parts')[0]['parts']
    for i in parts_one:
        parts_two = i['parts']
        for i in parts_two:
            data = i['body']['data']
            break
        break
    data = data.replace("-", "+").replace("_", "/")
    decoded_data = base64.b64decode(data)
    soup = BeautifulSoup(decoded_data, "lxml")
    body = soup.body()
    text = soup.find('p').text.strip().splitlines()
    slist = list(filter(None, text))
    text_check = 0
    text = ''
    for i in slist:
        try:
            if text_check == 1:
                text = text + i
            else:
                raise 'no'
        except:
            if 'The response' in i:
                text_check = 1
            continue
    return text


def failure_notice(payload):
    data = payload['body']['data']
    data = data.replace("-", "+").replace("_", "/")
    decoded_data = base64.b64decode(data).decode()
    f = decoded_data.find('Bcc:')
    failed_email = re.findall(r'[\w.+-]+@[\w-]+\.[\w.-]+', decoded_data[f:])[0]
    return failed_email


def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
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
        filename = input('Enter File Name: ') + '.xlsx'
        service = build('gmail', 'v1', credentials=creds, static_discovery=False)
        user = service.users().getProfile(userId='me').execute()
        userId = user['emailAddress']
        main_df = pd.DataFrame()
        email_no = 0
        check_totol = False
        total_email_want = int(input("Enter How Many Bounced Email You Want: "))
        results = service.users().messages().list(userId='me', labelIds='INBOX').execute()
        while True:
            for i in results['messages']:
                run_data = False
                txt = service.users().messages().get(userId='me', id=i['id']).execute()
                payload = txt['payload']
                headers = payload['headers']
                subject = None
                description = None
                failed_email = None
                for one in headers:
                    if one['name']=='Subject' and 'failure notice' in one['value']:
                        run_data = True
                        subject = one['value']
                        failed_email = failure_notice(payload)
                        try:
                            description = text(payload)
                        except:
                            description = 'None'
                        email_no += 1
                        break
                    if one['name'] == 'From' and 'Mail Delivery' in one['value'] or one['name'] == 'Subject' and 'Delivery Status Notification'  in one['value']:
                        run_data = True
                        for three in headers:
                            if three['name'] == 'Subject':
                                subject = three['value']
                                break
                        try:
                            failed_email = mail_delivery_1(payload)
                        except:
                            failed_email = mail_delivery_2(payload)
                        try:
                            description = text(payload)
                        except:
                            description = 'None'
                
                        email_no += 1
                        break
                    # else:
                    #     pass
                    try:
                        if one['name'] == 'Subject' and 'Undeliverable' in one['value'] or one['name'] == 'From' and 'postmaster' in one['value']:
                            run_data = True
                            for i in headers:
                                if i['name'] == 'Subject':
                                    subject = i['value']
                            failed_email = postmaster_1(payload)
                            email_no += 1
                            break
                    except:
                        pass
                if run_data is True:
                    if failed_email == userId or failed_email is None:
                        try:
                            failed_email = full(payload)
                        except:
                            failed_email = None
                    else:
                        pass
                    if failed_email is None:
                        try:
                            failed_email = postmaster_2(txt)
                        except:
                            pass
                    else:
                        pass
                    df = pd.DataFrame({'Bounced email': [failed_email],
                                       'Subject': [subject],
                                       'Description': [description]})
                    main_df = pd.concat([main_df, df])
                    percent = email_no / total_email_want * 100
                    print(percent)
                    if email_no >= total_email_want:
                        check_totol = True
                        break
            if check_totol is True:
                break
            else:
                try:
                    nextpagetoken = results['nextPageToken']
                    results = service.users().messages().list(userId='me', labelIds='INBOX',
                                                              pageToken=nextpagetoken).execute()
                except:
                    print('Your Total Bounced Email is ', email_no)
                    break
        try:
            main_df.to_excel(filename,index=False)
            print('Completed')
            # return None
        except Exception as e:
            print(e)            

    except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
        print(f'An error occurred: {error}')


main()