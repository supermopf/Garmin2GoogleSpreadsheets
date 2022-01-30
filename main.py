#!/usr/bin/env python3
import logging
import os.path
import pickle
from datetime import date, datetime

import google.auth.transport.requests
import google_auth_oauthlib.flow
import googleapiclient.discovery
from garminconnect import (
    Garmin,
    GarminConnectConnectionError,
    GarminConnectTooManyRequestsError,
    GarminConnectAuthenticationError,
)

import config

logging.basicConfig(level=logging.ERROR)

today = date.today()

try:
    client = Garmin(config.garmin["USERNAME"], config.garmin["PASSWORD"])
except (
        GarminConnectConnectionError,
        GarminConnectAuthenticationError,
        GarminConnectTooManyRequestsError,
) as err:
    print("Error occurred during Garmin Connect Client init: %s" % err)
    quit()
except Exception:  # pylint: disable=broad-except
    print("Unknown error occurred during Garmin Connect Client init")
    quit()

try:
    client.login()
except (
        GarminConnectConnectionError,
        GarminConnectAuthenticationError,
        GarminConnectTooManyRequestsError,
) as err:
    print("Error occurred during Garmin Connect Client login: %s" % err)
    quit()
except Exception:  # pylint: disable=broad-except
    print("Unknown error occurred during Garmin Connect Client login")
    quit()

weightList = client.get_body_composition('2000-01-01', today.isoformat())
weightList = weightList["dateWeightList"]

result = []

for weightListItem in weightList:
    if weightListItem['boneMass'] is not None:
        row = (
            datetime.fromtimestamp(weightListItem['date'] / 1e3).strftime("%d.%m.%Y"),
            (weightListItem['weight'] / 1000),
            weightListItem['bmi'],
            weightListItem['bodyFat'],
            weightListItem['bodyWater'],
            (weightListItem['boneMass'] / 1000),
            (weightListItem['muscleMass'] / 1000),
        )
        result.append(row)

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
        creds.refresh(google.auth.transport.requests.Request())
    else:
        flow = google_auth_oauthlib.flow.InstalledAppFlow.from_client_secrets_file(
            'credentials.json', config.spreadsheet['SCOPES'])
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = googleapiclient.discovery.build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()

value_range_body = {
    "range": config.spreadsheet["RANGE_NAME"],
    "majorDimension": 'ROWS',
    "values": result
}
request = sheet.values().update(spreadsheetId=config.spreadsheet['SPREADSHEET_ID'],
                                range=config.spreadsheet["RANGE_NAME"],
                                valueInputOption="USER_ENTERED",
                                body=value_range_body)
response = request.execute()
