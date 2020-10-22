#!/usr/bin/env python3
import time

from garminconnect import (
    Garmin,
    GarminConnectConnectionError,
    GarminConnectTooManyRequestsError,
    GarminConnectAuthenticationError,
)
import config
import pickle
import os.path
import googleapiclient.discovery
import google_auth_oauthlib.flow
import google.auth.transport.requests

from datetime import date

import logging
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

weightList = client.fetch_data(
    "https://connect.garmin.com/modern/proxy/weight-service/weight/daterangesnapshot?startDate=2000-01-01&endDate=" + today.isoformat())
weightList = weightList["dateWeightList"]

result = []

for weightListItem in weightList:
    if weightListItem['boneMass'] is not None:
        row = (
            time.strptime(weightListItem['date'], "%x"),
            (weightListItem['weight'] / 1000),
            weightListItem['bmi'],
            weightListItem['bodyFat'],
            weightListItem['bodyWater'],
            (weightListItem['boneMass'] / 1000),
            (weightListItem['muscleMass'] / 1000),
        )
        result.append(row)

print(result)

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
    "range": "Import!A2:G",
    "majorDimension": 'ROWS',
    "values": result
}
request = sheet.values().update(spreadsheetId=config.spreadsheet['SPREADSHEET_ID'],
                                range="Import!A2:G",
                                valueInputOption="USER_ENTERED",
                                body=value_range_body)
response = request.execute()
