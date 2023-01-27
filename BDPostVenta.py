import pyodbc
import pandas as pd
import numpy as np
from Google import create_service

##conexión a BD
DRIVER_NAME ='SQL SERVER'
SERVER_NAME = '192.168.50.201'
DATABASE_NAME = 'pv'
UID='consultapluto'
PWD='_pepe@2015'; 



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
cursor.execute("SELECT     dbo.V_DETALLE_RECLAMOS_OT.id_reclamo, dbo.V_DETALLE_RECLAMOS_OT.id_ot, dbo.RECLAMOS.id_propiedad, dbo.V_DETALLE_RECLAMOS_OT.nombre_obra,                       dbo.V_DETALLE_RECLAMOS_OT.nombre_etapa, dbo.V_DETALLE_RECLAMOS_OT.direccion, dbo.V_DETALLE_RECLAMOS_OT.lote, dbo.V_DETALLE_RECLAMOS_OT.manzana,                       dbo.V_DETALLE_RECLAMOS_OT.nombres, dbo.V_DETALLE_RECLAMOS_OT.estado, dbo.V_DETALLE_RECLAMOS_OT.recinto, dbo.V_DETALLE_RECLAMOS_OT.lugar,                       dbo.V_DETALLE_RECLAMOS_OT.item, dbo.V_DETALLE_RECLAMOS_OT.problema, CAST(dbo.V_DETALLE_RECLAMOS_OT.fecha_reclamo AS DATE)as Fecha_Reclamo, CAST(dbo.V_DETALLE_RECLAMOS_OT.fecha_conformiad as DATE)as Fecha_Conformidad,                       dbo.V_DETALLE_RECLAMOS_OT.conformidad, dbo.MAESTRO_ESTADOS.descripcion, dbo.RECHAZOCALIDAD.DESCRIPCION AS Expr1, dbo.MOTIVOPOSTERGADO.POSTERGADO, dbo.V_DETALLE_RECLAMOS_OT.telefono1 FROM         dbo.V_DETALLE_RECLAMOS_OT INNER JOIN                      dbo.RECLAMOS ON dbo.V_DETALLE_RECLAMOS_OT.id_reclamo = dbo.RECLAMOS.id_reclamo INNER JOIN                      dbo.ORDENES_TRABAJO ON dbo.V_DETALLE_RECLAMOS_OT.id_ot = dbo.ORDENES_TRABAJO.id_ot INNER JOIN                      dbo.MAESTRO_ESTADOS ON dbo.RECLAMOS.estado = dbo.MAESTRO_ESTADOS.codigo_estado INNER JOIN                      dbo.RECHAZOCALIDAD ON dbo.ORDENES_TRABAJO.rechazocalidad = dbo.RECHAZOCALIDAD.id INNER JOIN                      dbo.MOTIVOPOSTERGADO ON dbo.ORDENES_TRABAJO.motivopostergado = dbo.MOTIVOPOSTERGADO.id Where dbo.V_DETALLE_RECLAMOS_OT.fecha_reclamo >= '2021-1-10'")
data=cursor.fetchall()
data_PV=pd.DataFrame(np.array(data)) 

"""
Getting  Google Sheets
"""
CLIENT_SECRET_FILE = 'client_secret.json'
API_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

service = create_service(CLIENT_SECRET_FILE, API_NAME, API_VERSION,SCOPES )

google_sheets_id = '1ZcOciAFVD66ll0hcACAGaLLueu01ofQGXCLqzLce9RQ'

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
    range='SABANA!A2:V'
    ).execute()


recordset = data_PV.values.tolist()

"""
Insert rows
"""
request_body_values = construct_request_body(recordset)
service.spreadsheets().values().clear(spreadsheetId=google_sheets_id, range='SABANA!A2:V').execute()
service.spreadsheets().values().update(
    spreadsheetId=google_sheets_id,
    valueInputOption='USER_ENTERED',
    range='SABANA!A2:V',
    body=request_body_values
    ).execute()

print('Task is complete')

cursor.close()