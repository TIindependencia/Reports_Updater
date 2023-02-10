import pyodbc
import pandas as pd
import numpy as np
from Google import create_service
from datetime import datetime, timezone
from pytz import timezone

##conexión a BD
DRIVER_NAME = 'ODBC Driver 17 for SQL Server'
SERVER_NAME = '192.168.1.31,1433'
DATABASE_NAME = 'DynamicsAxProd'
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


"""
Getting  Google Sheets
"""
CLIENT_SECRET_FILE = 'client_secret.json'
API_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

service = create_service(CLIENT_SECRET_FILE, API_NAME, API_VERSION,SCOPES )

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


def get_values(spreadsheet_id,range):
    request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range,majorDimension='ROWS')
    response = request.execute()
    values = response['values'][0][0]
    return values

def load_values(df,id,range):
    request_body_values=construct_request_body(df.values.tolist())
    service.spreadsheets().values().append(
            spreadsheetId=id,
            valueInputOption='USER_ENTERED',
            range=range,
            body=request_body_values
        ).execute()

def reset_values(id):
    service.spreadsheets().values().clear(spreadsheetId=id, range='Empresas AX1!A2:V').execute()

cursor=conn.cursor()
cursor.execute("SELECT [DATAAREAID]      ,[AMOUNTMSTC]      ,[AMOUNTMSTD]      ,[ACCOUNTNUM]      ,[CLOSED]      ,[DOCUMENTCODE]      ,[DUEDATE]      ,[INVOICE]      ,[POSTINGPROFILE]      ,[REPORTINGCURRENCYAMOUNT]      ,[AMOUNTMSTR]      ,[TRANSDATE]      ,[TRANSTYPE]      ,[TXT]      ,[VOUCHER]      ,[NAME]      ,[AMOUNTCURTR]      ,[CURRENCYCODE]      ,[APPROVED], [dimCostCenter]  FROM [GoogleDrive].[dbo].[VENDTRANSRESIDUE] where DATAAREAID<>'inds' ")
data=cursor.fetchall()
df=pd.DataFrame(np.array(data))

id='1GfeaFnoMcZAUXhkj_EFRqvet6efJkzcaBY06R43Dda8'
df[4]=df[4].astype(str)
df[6]=df[6].astype(str)
df[11]=df[11].astype(str)

now_utc = datetime.now(timezone('UTC'))
now = now_utc.astimezone(timezone('America/Santiago'))

try:
    reset_values(id)
    load_values(df,id,'Empresas AX1!A2:V')
    print("Data Actualizada "+ str(now))
except Exception as e:
    print(e)
    print(now)


    