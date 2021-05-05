import pickle
import os
from datetime import datetime
from google_auth_oauthlib.flow import Flow, InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request


class NewCalendar:
    def make_new_calendar(summary='calendarSummary', time_zone='Asia/Jerusalem'):
        service = make_service()
        calendar = {
            'summary': summary,
            'timeZone': time_zone
        }
        created_calendar = service.calendars().insert(body=calendar).execute()
        print(created_calendar['id'])
        return created_calendar['id']


def make_service(client_secret_file="credentials.json", api_name='calendar', api_version='v3',
                 scopes=['https://www.googleapis.com/auth/calendar']):
    indicator = os.path.exists(client_secret_file)
    while not indicator:
        print("you dont have client_secret_file that Authentication you enter to your mail.")
        print("alternatively you can make new mail and to export the created calendar to your official"
              " account afterwords")
        print(" to make new gmail account press:"
              " 'https://accounts.google.com/SignUp?service=mail&continue=https://mail.google.com/mail/' ")
        print(" for create client_secret_file by google site press:"
              "'https://console.cloud.google.com/apis/credentials/oauthclient?previousPage=%2Fapis%2Fcredentials"
              "%3Fpli%3D1%26project%3Dhebrew-calender--1618848664482&project=hebrew-calender--1618848664482' "
              "or search in google fot 'client_secret_file")
        print("enter here the path to the client_secret_file you generate locally")

        client_secret_file = input()
        indicator = os.path.exists(client_secret_file)

    # todo need to change the below content the personal user google API.
    # generated from:  https://developers.google.com/workspace/guides/create-credentials?hl=en
    service = create_service(client_secret_file, api_name, api_version, scopes)
    return service


def create_service(client_secret_file, api_name, api_version, *scopes):
    print(client_secret_file, api_name, api_version, scopes, sep='-')
    client_secret_file = client_secret_file
    api_service_name = api_name
    api_version = api_version
    scopes = [scope for scope in scopes[0]]
    print(scopes)

    cred = None

    pickle_file = f'token_{api_service_name}_{api_version}.pickle'
    # print(pickle_file)

    if os.path.exists(pickle_file):
        with open(pickle_file, 'rb') as token:
            cred = pickle.load(token)

    if not cred or not cred.valid:
        if cred and cred.expired and cred.refresh_token:
            cred.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(client_secret_file, scopes)
            cred = flow.run_local_server()

        with open(pickle_file, 'wb') as token:
            pickle.dump(cred, token)

    try:
        service = build(api_service_name, api_version, credentials=cred)
        print(api_service_name, 'service created successfully')
        return service
    except Exception as e:
        print('Unable to connect.')
        print(e)
        return None


def convert_to_RFC_datetime(year=1900, month=1, day=1, hour=0, minute=0):
    dt = datetime.datetime(year, month, day, hour, minute, 0).isoformat() + 'Z'
    return dt
