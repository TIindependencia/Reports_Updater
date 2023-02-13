import pyodbc
import pandas as pd
import numpy as np
from Google import create_service
from datetime import datetime
from pytz import timezone

DRIVER_NAME = 'ODBC Driver 17 for SQL Server'
SERVER_NAME = '192.168.1.31,1433'
DATABASE_NAME = 'pv'
UID='google'
PWD='Pago1010.' 

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

id='1GSCfOjTOhfab8e-bb9pcWCNNYFJFUOSlS5fxj1GBDDg'

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

cursor=conn.cursor()
cursor.execute("SELECT * FROM dbo.V_INFORME_DETALLE_VIVIENDA_RECLAMOOT_ACTIVA")
data=cursor.fetchall()

df=pd.DataFrame(np.array(data))
del df[13]
del df[14]

service.spreadsheets().values().clear(spreadsheetId=id, range='DETALLE!A2:L').execute()
now_utc = datetime.now(timezone('UTC'))
now = now_utc.astimezone(timezone('America/Santiago'))

request_body_values=construct_request_body(df.values.tolist())

try:
    service.spreadsheets().values().append(
            spreadsheetId=id,
            valueInputOption='USER_ENTERED',
            range='DETALLE!A2:L',
            body=request_body_values
        ).execute()
    print("Data Actualizada "+ str(now))
except Exception as e:
    print(e)
    print(now)

