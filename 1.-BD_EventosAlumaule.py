import pyodbc
import pandas as pd
import numpy as np
from Google import create_service
from datetime import datetime
from pytz import timezone


##conexión a BD
DRIVER_NAME = 'ODBC Driver 17 for SQL Server'
SERVER_NAME = '192.168.1.31,1433'
DATABASE_NAME = 'pv'
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

cursor = conn.cursor()
cursor.execute("SELECT * FROM dbo.V_INFORME_DETALLE_VIVIENDA_RECLAMOOT_ACTIVA")
data=cursor.fetchall()
data_PV=pd.DataFrame(np.array(data)) 
data_PV.pop(data_PV.columns[-1])
data_PV.pop(data_PV.columns[-1])
"""
Getting  Google Sheets
"""
CLIENT_SECRET_FILE = 'client_secret.json'
API_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

service = create_service(CLIENT_SECRET_FILE, API_NAME, API_VERSION,SCOPES )

google_sheets_id = '1QylL7IcQn825ycs_qMcnP30tJuY6YK64_epow64q1eM'

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


now_utc = datetime.now(timezone('UTC'))
now = now_utc.astimezone(timezone('America/Santiago'))
print(now)
recordset = data_PV.values.tolist()

"""
Insert rows
"""
request_body_values = construct_request_body(recordset)
service.spreadsheets().values().clear(spreadsheetId=google_sheets_id, range='BD!A2:M').execute()
service.spreadsheets().values().update(
    spreadsheetId=google_sheets_id,
    valueInputOption='USER_ENTERED',
    range='BD!A2:M',
    body=request_body_values
    ).execute()

print('Task is complete')

cursor.close()