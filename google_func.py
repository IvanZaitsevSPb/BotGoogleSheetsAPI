# System Imports
import os
import os.path

# Data Type Imports
import pickle
import psutil
import time
import pandas as pd
import re

# Imports for working with Google Sheets
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


# Measuring the memory and running time of the program
process = psutil.Process()
start_time = time.time()


# Class for one-time authorization, saves 'token.pickle', defines column creation and table filling
class GoogleSheet:
    SPREADSHEET_ID = ''
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    service = None

    def __init__(self):
        creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                print('flow')
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', self.SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        self.service = build('sheets', 'v4', credentials=creds)

    def updateRangeValues(self, range, values):
        data = [{
            'range': range,
            'majorDimension': 'ROWS',
            'values': values
        }]
        body = {
            'valueInputOption': 'USER_ENTERED',
            'data': data
        }
        result = self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.SPREADSHEET_ID,
                                                                  body=body).execute()
        print('{0} cells updated.'.format(result.get('totalUpdatedCells')))

    def insertColumn(self, startIndex):
        SHEET_ID = '1733389309'
        request_body = {
            "requests": [
                {
                    "insertDimension": {
                        "range": {
                            "sheetId": SHEET_ID,
                            "dimension": "COLUMNS",
                            "startIndex": startIndex,
                            "endIndex": startIndex + 1
                        },
                        "inheritFromBefore": True
                    }
                }
            ]
        }
        self.service.spreadsheets().batchUpdate(spreadsheetId=self.SPREADSHEET_ID,
                                                body=request_body).execute()


# Function for reading and converting data from xlsx to Goggle Sheets filling format
def convert_xlsx(name_table):

    search_date = pd.read_excel(name_table)
    search_date1 = search_date.iloc[5:6:].fillna(' ').to_string(header=False, index=False)
    dates = re.findall(r'\d{2}.\d{2}.\d{4}', search_date1)
    names_column = ['Номер прибора учета', 'маршрут', 'Физический адрес', 'Уровень доступа DLMS', 'Пароль']
    names_column.extend(dates)
    database = pd.read_excel(name_table, header=6, usecols=names_column)
    single_column = database.fillna(' ')
    convert_list = [list(x) for x in single_column.to_records(index=False)]
    convert_list.insert(0, names_column)
    return convert_list


# Function to create a common summary data table, we take data from it and push it to Google
def create_xlsx_from_list(data_list, file_name):
    df = pd.DataFrame(data_list)
    df.to_excel(file_name, index=False)


# Function for pushing the main xlsx table
def prepare_for_push(name_table):
    search_date = pd.read_excel(name_table)
    del_line = search_date.iloc[0::].fillna(' ')
    prepare_list = [list(x) for x in del_line.to_records(index=False)]
    return prepare_list


# main func
def main_func(list_names):
    list_data = list()
    gs = GoogleSheet()
    gs.insertColumn(14)  # Enter the number of the column that we are creating
    for i in range(len(list_names)):
        if len(list_data) == 0:
            list_data.extend(convert_xlsx(list_names[i]))
        else:
            list_data.extend(convert_xlsx(list_names[i])[1:])
    create_xlsx_from_list(list_data, 'test.xlsx')
    range_data = ''  # enter the filled range of the table
    prepare = prepare_for_push('test.xlsx') # take the data from the pivot table
    gs.updateRangeValues(range_data, prepare)
    print("--- %s seconds ---" % (time.time() - start_time))  # Working hours of the program
    print(process.memory_info().rss)  # Memory consumed in bytes


