from googleapiclient.discovery import build
import os
import pickle as pkl


def get_google_sheet(SCOPES, SAMPLE_SPREADSHEET_ID, SAMPLE_RANGE_NAME):
    '''
    將指定輸入的工作表欄位輸出
    '''
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pkl.load(token)
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])
    return values
