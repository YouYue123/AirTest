from __future__ import print_function
import pygraphviz as pgv
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import os.path

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1uPvwCkQmBNqdQOhmoTBEHWrK6aaqiRc4EQa2QVHkiY4'

dict = {}
def readData(data):
    for item in data[1:]:
        dict[item[0]] = {"parent": item[3],
                         "description": item[1], 
                         "flow": item[2],
                        "status": item[-1]
                        }


def main():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().batchGet(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                     ranges=["Project", "VirtualSite", "Progress dashboard", 
                                             "Link", "Markup dashboard", "PDF Report", 
                                             "Videowalk", "Attachment in VirtualSite", "Markup in VirtualSite", "BIM", 
                                             "Labels and categories", "Usage", "Project Settings", "Miscellaneous"]).execute()
    list = result["valueRanges"]
    for item in list:
        data = item["values"]
        readData(data)
    passedCases = 0
    failedCases = 0
    unTestedCases = 0
    for key in dict.keys():
        item = dict[key]
        if item["status"] == "P":
          passedCases += 1
        elif item["status"] == "F":
          failedCases += 1
        else:
          unTestedCases += 1
    print(f'Tested cases {passedCases + failedCases + unTestedCases}')
    print(f'Test coverage {((passedCases + failedCases) / (passedCases + failedCases + unTestedCases)) * 100}%')
    print(f'Successful: {passedCases} Failed: {failedCases}')
    print("Done!")


if __name__ == "__main__":
    main()
