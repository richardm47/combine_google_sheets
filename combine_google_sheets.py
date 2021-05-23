import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import pandas as pd

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

def extract_sheet_data(sheet_id):
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
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
            flow = InstalledAppFlow.from_client_secrets_file('gartner.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    sheet_metadata = sheet.get(spreadsheetId=sheet_id).execute() ##extract spreadsheet metadata to get sheet names
    sheet_names = [i['properties']['title'] for i in sheet_metadata.get('sheets')] ##names of all sheets within spreadsheet

    doc_df = []
    for sheet_name in sheet_names:
        result = sheet.values().get(spreadsheetId=sheet_id,
                                    range=sheet_name).execute()
        values = result.get('values', [])
        
        df = pd.DataFrame(data=values[1:], columns=values[0]) ## create pandas df from sheet data
        doc_df.append(df)

    return pd.concat(doc_df, sort=True, ignore_index=True) ##merge multiple sheets data to single sheet



if __name__ == '__main__':
    
    ##links of google sheets with read permission enabled
    gsheets = ["https://docs.google.com/spreadsheets/d/1jOtdqMaqCQj3NtGmatvx3Pc4pbfglV0xs2bbu4JdWn0/edit?usp=sharing",
               "https://docs.google.com/spreadsheets/d/1mLUVzCvKHYfbTJxdXzPRAVs2AO9Tktw5ewwow3LImzU/edit?usp=sharing",
               "https://docs.google.com/spreadsheets/d/1MYVH6LaChZ23A1Bk6OBKueP8WqIpjigjjPfTwcXCcH0/edit?usp=sharing"]

    sheet_ids = [gsheet.split("/")[5] for gsheet in gsheets]  ## fetch google spreadsheet IDs

    all_dfs = []
    for sheet_id in sheet_ids:
        df = extract_sheet_data(sheet_id) #fetch data of google each spreadsheet as pandas df
        all_dfs.append(df)

    combined_df = pd.concat(all_dfs, sort=True, ignore_index=True)  ##merge multiple spreadsheet data to single spreadsheet
    combined_df.to_excel("output.xlsx") ##save data to excel