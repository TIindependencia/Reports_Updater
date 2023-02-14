import pyodbc
import pandas as pd
import numpy as np
from Google import create_service
from datetime import datetime
from pytz import timezone


##conexión a BD
DRIVER_NAME ='ODBC Driver 17 for SQL Server'
SERVER_NAME = '192.168.1.31,1433'
DATABASE_NAME = 'GoogleDrive'
UID='google'
PWD='Pago1010.'; 

try:
    conn = pyodbc.connect('DRIVER='+DRIVER_NAME+';SERVER='+SERVER_NAME+';DATABASE='+DATABASE_NAME+';ENCRYPT=no;UID='+UID+';PWD='+ PWD +'')
    print('Connection created') 
except pyodbc.DatabaseError as e:
    print('Database Error 1:')
    print(str(e.value[1]))
except pyodbc.Error as e:
    print('Connection Error 2:')
    print(str(e.value[1]))

cursor=conn.cursor()
cursor.execute('SELECT * FROM CI_VCOSAPPROVEDRIVE')
data=cursor.fetchall()
df=pd.DataFrame(np.array(data))

df[11]=df[11].astype(str)
df[12]=df[12].astype(str)
df[17]=df[17].astype(str)
df[21]=df[21].astype(str)
df[25]=df[25].astype(str)

"""
Getting  Google Sheets
"""
CLIENT_SECRET_FILE = 'client_secret.json'
API_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

service = create_service(CLIENT_SECRET_FILE, API_NAME, API_VERSION,SCOPES )

google_sheets_id = '1h5dOstXMJukXV7o5Su1oq0hNRIkas0ndvsrEKHbHE5k'

def construct_request_body(value_array, dimension: str='ROWS') -> dict:
    try:
        request_body = {
            'majorDimension': dimension,
            'values': value_array
        }
        return request_body
    except Exception as e:
        print(e)
        return {}

response = service.spreadsheets().values().get(
    spreadsheetId=google_sheets_id,
    majorDimension='ROWS',
    range='BD2!A2:AC'
    ).execute()


recordset = df.values.tolist()

"""
Insert rows
"""
request_body_values = construct_request_body(recordset)
service.spreadsheets().values().clear(spreadsheetId=google_sheets_id, range='BD2!A2:AC').execute()
now_utc = datetime.now(timezone('UTC'))
now = now_utc.astimezone(timezone('America/Santiago'))

try:
    service.spreadsheets().values().update(
        spreadsheetId=google_sheets_id,
        valueInputOption='USER_ENTERED',
        range='BD2!A2:AC',
        body=request_body_values
        ).execute()
    cursor.close()
    print('Task is complete')
    print(now)
except Exception as e:
    print(e)
    print (now)



