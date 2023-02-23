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

def getValues(range):
    array=[]
    request = service.spreadsheets().values().get(spreadsheetId=google_sheets_id, range=range)
    response = request.execute()
    for x in response["values"]:
        if len(x)!=0:
            array.append(x[0])
    return array

def getStatement(array):
    stmt=''
    for centro in array:
        aux="PROJID='"+centro+"' OR "
        stmt=stmt+aux
    stmt=stmt[:-3]
    return stmt

def getStatementC(array):
    stmt=''
    for centro in array:
        aux="CI_PURCHASEMANAGEMENTLINE.PROJID='"+centro+"' OR "
        stmt=stmt+aux
    stmt=stmt[:-3]
    return stmt

values=getValues('AUX!C2:C')
where_stmt=getStatement(values)

cursor=conn.cursor()
stmt_gestion='SELECT PROJID AS PROYECTO		,PROJNAME AS NOMBRE_PROYECTO				,ITEMID AS CODIGO_ARTICULO		,ITEMNAME AS NOMBRE_ARTICULO		,UNITID AS UNIDAD		,cast(AMOUNT7 as int) as CANTIDAD_ORIGINAL		,cast(PURCHPRICE as int) as PRECIO_COMPRA		,cast(QTY as int) AS CANTIDAD_MODIFICADA		,cast(PRICE as int) AS UNITARIO_PRESUPUESTADO		,cast(AMOUNT6 as int) AS ORIGINAL		,cast(AMOUNT as int) AS PRESUPUESTO		,cast(PURCHQTY as int) AS CANTIDAD_COMPRA		,cast(INPUTSQTY as int) AS RECIBIDO 		,cast(AMOUNT1 as int) AS COMPRADO		,cast(AMOUNT2 as int) AS DISPONIBLE		,cast(AMOUNT3_5 as int) AS CANTIDAD_POR_COMPRAR		,cast(AMOUNT3 as int) AS POR_COMPRAR		,cast(AMOUNT4 as int)AS PROYECTADO		,cast(AMOUNT5 as int) AS AHORRO_PERDIDA		,CATEGORYID AS CATEGORIA		,CATEGORYNAME1 AS NOMBRE_CATEGORIA 		,DATAAREAID AS EMPRESA, ITEMID AS ITEMID  FROM CI_PURCHASEMANAGEMENTLINE WHERE '
cursor.execute(stmt_gestion+where_stmt)
data=cursor.fetchall()
Gestion=pd.DataFrame(np.array(data))

stmt_detalle='Select PURCHQTY AS CANTIDAD , RECEIVED AS RECEPCIONADA, LINEAMOUNT AS IMPORTENETO, PURCHID AS ORDENDECOMPRA, PURCHPRICE AS UNITARIO, REMAINPURCHPHYSICAL AS PORRECEPCIONAR, (PURCHPRICE*REMAINPURCHPHYSICAL) PORRECPD,PURCHNAME AS PROVEEDOR, ITEMID AS ITEMID, PROJID AS COD_PROYECTO FROM PURCHDETAILITEMDRIVE WHERE '
cursor.execute(stmt_detalle+where_stmt)
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

stmt_where_consolidado=getStatementC(values)
cursor.execute("SELECT CI_PURCHASEMANAGEMENTLINE.PROJID AS PROYECTO ,CI_PURCHASEMANAGEMENTLINE.PROJNAME AS NOMBRE_PROYECTO ,CI_PURCHASEMANAGEMENTLINE.ITEMID AS CODIGO_ARTICULO ,CI_PURCHASEMANAGEMENTLINE.ITEMNAME AS NOMBRE_ARTICULO ,CI_PURCHASEMANAGEMENTLINE.UNITID AS UNIDAD , "
                + "cast(CI_PURCHASEMANAGEMENTLINE.AMOUNT7 as int) as CANTIDAD_ORIGINAL	,cast(CI_PURCHASEMANAGEMENTLINE.PURCHPRICE as int) as PRECIO_COMPRA ,cast(CI_PURCHASEMANAGEMENTLINE.QTY as int) AS CANTIDAD_MODIFICADA	"
                + ",cast(CI_PURCHASEMANAGEMENTLINE.PRICE as int) AS UNITARIO_PRESUPUESTADO ,cast(CI_PURCHASEMANAGEMENTLINE.AMOUNT6 as int) AS ORIGINAL	,cast(CI_PURCHASEMANAGEMENTLINE.AMOUNT as int) AS PRESUPUESTO	,cast(CI_PURCHASEMANAGEMENTLINE.PURCHQTY as int) AS CANTIDAD_COMPRA , "
                + "cast(CI_PURCHASEMANAGEMENTLINE.INPUTSQTY as int) AS RECIBIDO ,cast(CI_PURCHASEMANAGEMENTLINE.AMOUNT1 as int) AS COMPRADO	,cast(CI_PURCHASEMANAGEMENTLINE.AMOUNT2 as int) AS DISPONIBLE	,cast(CI_PURCHASEMANAGEMENTLINE.AMOUNT3_5 as int) AS CANTIDAD_POR_COMPRAR		, "
                + "cast(CI_PURCHASEMANAGEMENTLINE.AMOUNT3 as int) AS POR_COMPRAR		,cast(CI_PURCHASEMANAGEMENTLINE.AMOUNT4 as int)AS PROYECTADO	,cast(CI_PURCHASEMANAGEMENTLINE.AMOUNT5 as int) AS AHORRO_PERDIDA	,CI_PURCHASEMANAGEMENTLINE.CATEGORYID AS CATEGORIA, "
                + "CI_PURCHASEMANAGEMENTLINE.CATEGORYNAME1 AS NOMBRE_CATEGORIA  ,CI_PURCHASEMANAGEMENTLINE.DATAAREAID AS EMPRESA, CI_PURCHASEMANAGEMENTLINE.ITEMID AS ITEMID,PURCHDETAILITEMDRIVE.PURCHQTY AS CANTIDAD , " 
                + "PURCHDETAILITEMDRIVE.RECEIVED AS RECEPCIONADA, PURCHDETAILITEMDRIVE.LINEAMOUNT AS IMPORTENETO, PURCHDETAILITEMDRIVE.PURCHID AS ORDENDECOMPRA, PURCHDETAILITEMDRIVE.PURCHPRICE AS UNITARIO, PURCHDETAILITEMDRIVE.REMAINPURCHPHYSICAL AS PORRECEPCIONAR, (PURCHDETAILITEMDRIVE.PURCHPRICE*PURCHDETAILITEMDRIVE.REMAINPURCHPHYSICAL) AS PORRECPD,PURCHDETAILITEMDRIVE.PURCHNAME AS PROVEEDOR, PURCHDETAILITEMDRIVE.ITEMID AS ITEMID, PURCHDETAILITEMDRIVE.PROJID AS COD_PROYECTO "
                + "FROM CI_PURCHASEMANAGEMENTLINE " 
                + "LEFT JOIN  PURCHDETAILITEMDRIVE ON CI_PURCHASEMANAGEMENTLINE.PROJID = PURCHDETAILITEMDRIVE.PROJID AND CI_PURCHASEMANAGEMENTLINE.ITEMID =PURCHDETAILITEMDRIVE.ITEMID "
                + "WHERE "+ stmt_where_consolidado)
print(stmt_where_consolidado)
data=cursor.fetchall()
Consolidado=pd.DataFrame(np.array(data))

Consolidado[23]=Consolidado[23].astype(float)
Consolidado[24]=Consolidado[24].astype(float)
Consolidado[25]=Consolidado[25].astype(float)
Consolidado[27]=Consolidado[27].astype(float)
Consolidado[28]=Consolidado[28].astype(float)
Consolidado[29]=Consolidado[29].astype(float)
Consolidado.fillna(value='', inplace=True)

try:
    #inserta Gestion de compra
    upload_data(Gestion,'BD_GESTION!B2:X')
    #inserta el detalle
    upload_data(Detalle,'BD_DETALLE!B2:K')
    #inserta el consolidado
    upload_data(Consolidado,'Aux_Consolidado!A2:AG')
    cursor.close()
    print(now)
except Exception as e:
    print(e)
    print(now)