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

google_sheets_id = '1jbljGDHBbFCuTr1POjCszpMv3pvBQHehprEKH0nPRWs'

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
cursor.execute('SELECT PROJID AS PROYECTO		,PROJNAME AS NOMBRE_PROYECTO				,ITEMID AS CODIGO_ARTICULO		,ITEMNAME AS NOMBRE_ARTICULO		,UNITID AS UNIDAD		,cast(AMOUNT7 as int) as CANTIDAD_ORIGINAL		,cast(PURCHPRICE as int) as PRECIO_COMPRA		,cast(QTY as int) AS CANTIDAD_MODIFICADA		,cast(PRICE as int) AS UNITARIO_PRESUPUESTADO		,cast(AMOUNT6 as int) AS ORIGINAL		,cast(AMOUNT as int) AS PRESUPUESTO		,cast(PURCHQTY as int) AS CANTIDAD_COMPRA		,cast(INPUTSQTY as int) AS RECIBIDO 		,cast(AMOUNT1 as int) AS COMPRADO		,cast(AMOUNT2 as int) AS DISPONIBLE		,cast(AMOUNT3_5 as int) AS CANTIDAD_POR_COMPRAR		,cast(AMOUNT3 as int) AS POR_COMPRAR		,cast(AMOUNT4 as int)AS PROYECTADO		,cast(AMOUNT5 as int) AS AHORRO_PERDIDA		,CATEGORYID AS CATEGORIA		,CATEGORYNAME1 AS NOMBRE_CATEGORIA 		,DATAAREAID AS EMPRESA  FROM CI_PURCHASEMANAGEMENTLINE ')
data=cursor.fetchall()
Gestion=pd.DataFrame(np.array(data))

cursor.execute('Select PURCHQTY AS CANTIDAD , RECEIVED AS RECEPCIONADA, LINEAMOUNT AS IMPORTENETO, PURCHID AS ORDENDECOMPRA, PURCHPRICE AS UNITARIO, REMAINPURCHPHYSICAL AS PORRECEPCIONAR, (PURCHPRICE*REMAINPURCHPHYSICAL) PORRECPD,PURCHNAME AS PROVEEDOR FROM PURCHDETAILITEMDRIVE')
data=cursor.fetchall()
Detalle=pd.DataFrame(np.array(data))


Detalle[0]=Detalle[0].astype(float)
Detalle[1]=Detalle[1].astype(float)
Detalle[2]=Detalle[2].astype(float)
Detalle[4]=Detalle[4].astype(float)
Detalle[5]=Detalle[5].astype(float)
Detalle[6]=Detalle[6].astype(float)

now_utc = datetime.now(timezone('UTC'))
now = now_utc.astimezone(timezone('America/Santiago'))

try:
    #inserta Gestion de compra
    upload_data(Gestion,'BD_GESTION!A2:V')
    #inserta el detalle
    upload_data(Detalle,'BD_DETALLE!A2:H')
    cursor.close()
    print(now)
except Exception as e:
    print(e)
    print(now)