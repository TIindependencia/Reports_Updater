import pyodbc
import pandas as pd
import numpy as np
from Google import create_service
from datetime import datetime, timezone
from pytz import timezone

##conexión a BD
DRIVER_NAME = 'ODBC Driver 17 for SQL Server'
SERVER_NAME = '192.168.1.34,1433'
DATABASE_NAME = 'xci'
UID='consulta'
PWD='2021@cii'; 



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

id='1aUiONAONz2EskDCysgI6gljCEmS6XaDjdOVXL0EZIZc'

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
cursor.execute(("SELECT cod_documento, desc_documento, insertby, inserttime, user_id, rut_empresa, cod_tipo_documento, cod_estado_documento, cod_digitado, "
                              + "rut_contribuyente,num_documento, fecha_emision, fecha_recepcion_sii, monto_documento, num_seguimiento, cod_recepcionado, cod_en_erp, cod_centro_costo, "
                              + "cod_recepcion_dte, cod_oc_ok, monto_en_erp, monto_recepcion, cod_forma_pago,monto_neto, xml, fecha_libro, fecha_contabilizacion, cod_flag_acuse, cod_incluir, "
                              + "cod_estado_acuse_recibo, cod_referenciado, cod_cedido, cod_vb, cod_accion_pendiente, monto_exento, cod_revisar, num_guias_ref, ocs_ref, cod_precontabilizado, "
                              + "cod_cuenta_proveedor, fecha_vencimiento, saldo_documento, num_cuenta_contable, cod_tipo_compra FROM dbo.vista_por_contabilizar ORDER BY fecha_recepcion_sii"))
data=cursor.fetchall()
FacturasPorContabilizar=pd.DataFrame(np.array(data))
FacturasPorContabilizar[3]=FacturasPorContabilizar[3].astype(str)
FacturasPorContabilizar[11]=FacturasPorContabilizar[11].astype(str)
FacturasPorContabilizar[12]=FacturasPorContabilizar[12].astype(str)
FacturasPorContabilizar[25]=FacturasPorContabilizar[25].astype(str)
FacturasPorContabilizar[26]=FacturasPorContabilizar[26].astype(str)
FacturasPorContabilizar[40]=FacturasPorContabilizar[40].astype(str)

service.spreadsheets().values().clear(spreadsheetId=id, range='UNO!A2:AR').execute()
now_utc = datetime.now(timezone('UTC'))
now = now_utc.astimezone(timezone('America/Santiago'))

request_body_values=construct_request_body(FacturasPorContabilizar.values.tolist())

try:
    service.spreadsheets().values().append(
            spreadsheetId=id,
            valueInputOption='USER_ENTERED',
            range='UNO!A2:AR',
            body=request_body_values
        ).execute()
    print("Data Actualizada "+ str(now))
except Exception as e:
    print(e)
    print(now)


