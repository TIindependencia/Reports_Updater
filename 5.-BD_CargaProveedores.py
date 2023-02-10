import pyodbc
import pandas as pd
import numpy as np
from Google import create_service
from datetime import datetime
from pytz import timezone

indsIdGs='1Onvvh0ABn9KXqm06pdwW92TKtHlZFD8adQrU4ERT6x0'
crioIdGS='17fSc0pN_d0zBHrPHh1KQRVnICrmfHBTJD27CGN9_4GM'
colbIdGS='1KyL_DjuAqVqIamnrDVB_gR4OgQd49bUAgr7cSmLUoLc'
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

def load_values(df,id):
    request_body_values=construct_request_body(df.values.tolist())
    service.spreadsheets().values().append(
            spreadsheetId=id,
            valueInputOption='USER_ENTERED',
            range='DATOS!A2',
            body=request_body_values
        ).execute()

def reset_values(id):
    service.spreadsheets().values().clear(spreadsheetId=id, range='Datos!A2:R').execute()

cuentasInds=get_values(indsIdGs,'cuentas!A2')


cursor = conn.cursor()
cursor.execute("SELECT cast (ACCOUNTNUM as bigint),cast(AMOUNTCURTR as bigint),cast (AMOUNTMSTC as bigint), cast(AMOUNTMSTD as bigint), cast(AMOUNTMSTR as bigint),DOCUMENTCODE, case when convert(varchar (10),DOCUMENTDATE,101) = '01/01/1900' then '' else convert(varchar (10),DOCUMENTDATE,103) end,cast(DUEDATE as date),INVOICE,NAME,POSTINGPROFILE,cast(REPORTINGCURRENCYAMOUNT as bigint),cast (TRANSDATE as date),TRANSTYPE,TXT,VOUCHER,DATAAREAID,CREATED  FROM VENDTRANSREPORTTREASURY WHERE POSTINGPROFILE IN ("+cuentasInds+") AND DATAAREAID = 'inds' ")

ProveedoresInds=cursor.fetchall()
ProveedoresInds=pd.DataFrame(np.array(ProveedoresInds))
ProveedoresInds[0]=ProveedoresInds[0].astype(str)
ProveedoresInds[6]=ProveedoresInds[6].astype(str)
ProveedoresInds[7]=ProveedoresInds[7].astype(str)
ProveedoresInds[12]=ProveedoresInds[12].astype(str)

cuentasCrio=get_values(crioIdGS,'cuentas!A2')
cursor.execute("SELECT cast (ACCOUNTNUM as bigint),cast(AMOUNTCURTR as bigint),cast (AMOUNTMSTC as bigint), cast(AMOUNTMSTD as bigint), cast(AMOUNTMSTR as bigint),DOCUMENTCODE, case when convert(varchar (10),DOCUMENTDATE,101) = '01/01/1900' then '' else convert(varchar (10),DOCUMENTDATE,103) end,cast(DUEDATE as date),INVOICE,NAME,POSTINGPROFILE,cast(REPORTINGCURRENCYAMOUNT as bigint),cast (TRANSDATE as date),TRANSTYPE,TXT,VOUCHER,DATAAREAID,CREATED  FROM VENDTRANSREPORTTREASURY WHERE POSTINGPROFILE IN ("+cuentasCrio+") AND DATAAREAID = 'crio' ")

ProveedoresCrio=cursor.fetchall()
ProveedoresCrio=pd.DataFrame(np.array(ProveedoresCrio))
ProveedoresCrio[0]=ProveedoresCrio[0].astype(str)
ProveedoresCrio[6]=ProveedoresCrio[6].astype(str)
ProveedoresCrio[7]=ProveedoresCrio[7].astype(str)
ProveedoresCrio[12]=ProveedoresCrio[12].astype(str)

cuentasColb=get_values(colbIdGS,'cuentas!A2')
cursor.execute("SELECT cast (ACCOUNTNUM as bigint),cast(AMOUNTCURTR as bigint),cast (AMOUNTMSTC as bigint), cast(AMOUNTMSTD as bigint), cast(AMOUNTMSTR as bigint),DOCUMENTCODE, case when convert(varchar (10),DOCUMENTDATE,101) = '01/01/1900' then '' else convert(varchar (10),DOCUMENTDATE,103) end,cast(DUEDATE as date),INVOICE,NAME,POSTINGPROFILE,cast(REPORTINGCURRENCYAMOUNT as bigint),cast (TRANSDATE as date),TRANSTYPE,TXT,VOUCHER,DATAAREAID,CREATED  FROM VENDTRANSREPORTTREASURY WHERE POSTINGPROFILE IN ("+cuentasColb+") AND DATAAREAID = 'colb' ")

ProveedoresColb=cursor.fetchall()
ProveedoresColb=pd.DataFrame(np.array(ProveedoresColb))
ProveedoresColb[0]=ProveedoresColb[0].astype(str)
ProveedoresColb[6]=ProveedoresColb[6].astype(str)
ProveedoresColb[7]=ProveedoresColb[7].astype(str)
ProveedoresColb[12]=ProveedoresColb[12].astype(str)

now_utc = datetime.now(timezone('UTC'))
now = now_utc.astimezone(timezone('America/Santiago'))

try:
    reset_values(indsIdGs)
    load_values(ProveedoresInds,indsIdGs)
    reset_values(crioIdGS)
    load_values(ProveedoresCrio,crioIdGS)
    reset_values(colbIdGS)
    load_values(ProveedoresColb,colbIdGS)
    print("Data Actualizada "+ str(now))
except Exception as e:
    print(e)
    print(now)
