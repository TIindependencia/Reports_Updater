import pyodbc
import pandas as pd
import numpy as np
from Google import create_service
from datetime import datetime
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

google_sheets_id = '1cy-wqxXWBsisd8o_tqeVTzW9xLTGiPOVIlWOcYA0-8g'

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

def upload_data(data,sheetname):
    recordset = data.values.tolist()
    """
    Insert rows
    """
    request_body_values = construct_request_body(recordset)
    service.spreadsheets().values().clear(spreadsheetId=google_sheets_id, range=sheetname).execute()
    service.spreadsheets().values().update(
        spreadsheetId=google_sheets_id,
        valueInputOption='USER_ENTERED',
        range=sheetname,
        body=request_body_values
        ).execute()

    print(sheetname+' insertado completo')

cursor=conn.cursor()
cursor.execute('  SELECT [DATAAREAID],[ACCOUNTNUM],[NAME],[CHECKNUMBER],[TRANSDATE],[MATURITYDATE],[AMOUNTCURDEBIT],[CURRENCYCODE],[JOURNALNUM],[PDCSTATUS] FROM [DynamicsAxProd].[dbo].[CUSTVENDPDCREGISTERGDVIEW] ORDER BY TRANSDATE ASC')
data=cursor.fetchall()
cheques=pd.DataFrame(np.array(data))

cheques[4]=cheques[4].astype(str)
cheques[5]=cheques[5].astype(str)
cheques[6]=cheques[6].astype(float)

cheques[9].replace(2,'Contabilizado', inplace=True)
cheques[0].replace('crio','CONSTRUCTORA LA RIOJA SPA', inplace=True)
cheques[0].replace('inds','CONSTRUCTORA INDEPENDENCIA SPA', inplace=True)
cheques[0].replace('colb','CONSTRUCTORA COLBUN SPA', inplace=True)

now_utc = datetime.now(timezone('UTC'))
now = now_utc.astimezone(timezone('America/Santiago'))

try:
    #inserta el consolidado
    upload_data(cheques,'CHEQUES A FECHA FUTURA!A2:J')
    cursor.close()
    print(now)
except Exception as e:
    print(e)
    print(now)

