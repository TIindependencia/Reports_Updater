import pyodbc
import pandas as pd
import numpy as np
from Google import create_service

##conexión a BD
DRIVER_NAME = 'ODBC Driver 17 for SQL Server'
SERVER_NAME = '192.168.1.31,1433'
DATABASE_NAME = 'SistratosWeb'
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
cursor.execute("SELECT IDCONTRATO, cast( SUM(AnticipoValor)as int)  AS MontoComprado, cast(ISNULL((SELECT SUM(monto)  "
                 + "FROM IndCompraDescuento WHERE idanticipo = IndCompraContrato.IDCONTRATO "
                 + "GROUP BY idanticipo),0)as int) as MontoDescontado, cast((SUM(AnticipoValor) -ISNULL((SELECT SUM(monto)  "
                 + "FROM [IndCompraDescuento]WHERE idanticipo = IndCompraContrato.IDCONTRATO "
                 + "GROUP BY idanticipo),0))as int) as SaldoDescuento FROM IndCompraContrato GROUP BY IDCONTRATO")
ComprasDescto=cursor.fetchall()
ComprasDescto=pd.DataFrame(np.array(ComprasDescto)) 

cursor.execute("SELECT IDCONTRATO, cast(SUM(AnticipoValor)as int) as MontoAnticipado,cast(isnull((SELECT SUM(monto)  "
                 + "FROM IndAnticipoDescuento WHERE idanticipo = IndAnticiposContrato.IDCONTRATO "
                 + "GROUP BY idanticipo),0)as int) AS MontoDescontado, cast(SUM(AnticipoValor) -isnull((SELECT SUM(monto) "
                 + "FROM IndAnticipoDescuento WHERE idanticipo = IndAnticiposContrato.IDCONTRATO "
                 + "GROUP BY idanticipo),0)as int) as SaldoAnticipo FROM IndAnticiposContrato GROUP BY IDCONTRATO")
Anticipo=cursor.fetchall()
Anticipo=pd.DataFrame(np.array(Anticipo)) 

cursor.execute("SELECT CtoEmpresa, NombreEmpresa, Ctocodigo, NombreCCosto, RutContratista, NombreContratista, idcontrato, "
                 + "NumeroContrato, NombreContrato, cast(MontoAnticipo AS int), FechaAnticipo, cast(NumeroTraspaso AS int), "
                 + "AnticipoContabilizado, AnticipoLiquidado FROM IndAnticiposEEPP")
Historico=cursor.fetchall()
Historico=pd.DataFrame(np.array(Historico))

cursor.execute("SELECT TOP (100) PERCENT dbo.INDEPContratoCAB.CtoEmpresa as CodEmpresa,  dbo.maeEmpresa.empNombre as NombreEmpresa,dbo.INDEPContratoCAB.CtoCodigo "
                 + "as CodCentroCosto,dbo.INDMaeUNegocioActivas.CtoDescripcion as NombreCentroCosto,  dbo.INDEPContratoCAB.RutContratista, dbo.INDTrContratistas.RazonSocial, "
                 + "dbo.INDEPContratoCAB.IdContrato,cast(dbo.INDEPContratoCAB.NumeroContrato as int), dbo.INDEPContratoCAB.Descripcion, "
                 + "cast(dbo.INDEPContratoCAB.TotalContrato as int), cast(ISNULL ((SELECT SUM(TotalTratos) AS Expr1 FROM dbo.INDEPAvanceCAB "
                 + "WHERE (IdContrato = dbo.INDEPContratoCAB.IdContrato) and estado = 5), 0)as int) AS AvanceContrato   , "
                 + "cast(ROUND(ISNULL(dbo.INDEPContratoCAB.TotalContrato, 0) - ISNULL ((SELECT SUM(TotalTratos) AS Expr1 "
                 + "FROM dbo.INDEPAvanceCAB AS INDEPAvanceCAB_1  WHERE (IdContrato = dbo.INDEPContratoCAB.IdContrato) and estado = 5), 0),0)as int) AS "
                 + "SaldoContrato , cast(round(round( ISNULL(dbo.INDEPContratoCAB.TotalContrato, 0) - ISNULL((SELECT     SUM(INDEPAvanceCAB_2.TotalTratos) AS Expr1 "
                 + "FROM dbo.INDEPAvanceCAB AS INDEPAvanceCAB_2 WHERE (INDEPAvanceCAB_2.IdContrato = dbo.INDEPContratoCAB.IdContrato)), 0),0)* dbo.INDEPContratoCAB.RetencionEjecucion/100,0)as int) as Retencion, "
                 + "cast(ROUND((ISNULL(dbo.INDEPContratoCAB.TotalContrato, 0) - ISNULL ((SELECT SUM(TotalTratos) AS Expr1 "
                 + "FROM dbo.INDEPAvanceCAB AS INDEPAvanceCAB_1 WHERE (IdContrato = dbo.INDEPContratoCAB.IdContrato)), 0))- "
                 + "(round(round( ISNULL(dbo.INDEPContratoCAB.TotalContrato, 0) - ISNULL((SELECT SUM(INDEPAvanceCAB_2.TotalTratos) AS Expr1 FROM "
                 + "dbo.INDEPAvanceCAB AS INDEPAvanceCAB_2 WHERE (INDEPAvanceCAB_2.IdContrato = dbo.INDEPContratoCAB.IdContrato)), 0),0)*dbo.INDEPContratoCAB.RetencionEjecucion/100,0)),0)as int) as SaldoSinRetencion , "
                 + "cast((select isnull(sum(totaltrato),0) from indtrtratoscab where codesttrato='A'and  idcontrato=dbo.INDEPContratoCAB.IdContrato) + "
                 + "iSNULL  ((SELECT SUM(TotalTratos) FROM dbo.INDEPAvanceCAB WHERE IdContrato = dbo.INDEPContratoCAB.IdContrato And ( Estado is null or Estado=1)), 0 ) as int) as TratosNoPagados FROM dbo.INDEPContratoCAB "
                 +" INNER JOIN dbo.INDTrContratistas ON dbo.INDEPContratoCAB.RutContratista = dbo.INDTrContratistas.RutContratista "
                 + "INNER JOIN dbo.INDMaeUNegocioActivas ON dbo.INDEPContratoCAB.CtoEmpresa = dbo.INDMaeUNegocioActivas.CtoEmpresa AND dbo.INDEPContratoCAB.CtoCodigo = dbo.INDMaeUNegocioActivas.CtoCodigo "
                 + "INNER JOIN dbo.maeEmpresa ON dbo.INDMaeUNegocioActivas.CtoEmpresa = dbo.maeEmpresa.empCodigo WHERE (dbo.INDEPContratoCAB.AnoControl = '2020') "
                 + "ORDER BY dbo.INDEPContratoCAB.CtoEmpresa,dbo.INDEPContratoCAB.CtoCodigo,dbo.INDTrContratistas.RazonSocial"  )

BDoriginal=cursor.fetchall()
BDoriginal=pd.DataFrame(np.array(BDoriginal))

cursor.execute("SELECT TOP (100) PERCENT dbo.INDEPContratoCAB.CtoEmpresa as CodEmpresa,  dbo.maeEmpresa.empNombre as NombreEmpresa,dbo.INDEPContratoCAB.CtoCodigo as CodCentroCosto,dbo.INDMaeUNegocioActivas.CtoDescripcion as NombreCentroCosto,  dbo.INDEPContratoCAB.RutContratista, dbo.INDTrContratistas.RazonSocial,dbo.INDEPContratoCAB.IdContrato,cast(dbo.INDEPContratoCAB.NumeroContrato as int), dbo.INDEPContratoCAB.Descripcion,cast(dbo.INDEPContratoCAB.TotalContrato as int),cast(ISNULL ((SELECT SUM(TotalTratos) AS Expr1 FROM dbo.INDEPAvanceCAB WHERE (IdContrato = dbo.INDEPContratoCAB.IdContrato) and estado = 5), 0)as int) AS AvanceContrato, cast(ROUND(ISNULL(dbo.INDEPContratoCAB.TotalContrato, 0) - ISNULL ((SELECT SUM(TotalTratos) AS Expr1 FROM dbo.INDEPAvanceCAB AS INDEPAvanceCAB_1  WHERE (IdContrato = dbo.INDEPContratoCAB.IdContrato) and estado = 5), 0),0)as int) AS SaldoContrato , cast(round(round( ISNULL(dbo.INDEPContratoCAB.TotalContrato, 0) - ISNULL((SELECT     SUM(INDEPAvanceCAB_2.TotalTratos) AS Expr1  FROM dbo.INDEPAvanceCAB AS INDEPAvanceCAB_2 WHERE (INDEPAvanceCAB_2.IdContrato = dbo.INDEPContratoCAB.IdContrato)), 0),0)* dbo.INDEPContratoCAB.RetencionEjecucion/100,0)as int) as Retencion, cast(ROUND((ISNULL(dbo.INDEPContratoCAB.TotalContrato, 0) - ISNULL ((SELECT SUM(TotalTratos) AS Expr1  FROM dbo.INDEPAvanceCAB AS INDEPAvanceCAB_1 WHERE (IdContrato = dbo.INDEPContratoCAB.IdContrato)), 0))- (round(round( ISNULL(dbo.INDEPContratoCAB.TotalContrato, 0) - ISNULL((SELECT SUM(INDEPAvanceCAB_2.TotalTratos) AS Expr1 FROM  dbo.INDEPAvanceCAB AS INDEPAvanceCAB_2 WHERE (INDEPAvanceCAB_2.IdContrato = dbo.INDEPContratoCAB.IdContrato)), 0),0)*dbo.INDEPContratoCAB.RetencionEjecucion/100,0)),0)as int) as SaldoSinRetencion , cast((select isnull(sum(totaltrato),0) from indtrtratoscab where codesttrato='A'and  idcontrato=dbo.INDEPContratoCAB.IdContrato) +  iSNULL  ((SELECT SUM(TotalTratos) FROM dbo.INDEPAvanceCAB WHERE IdContrato = dbo.INDEPContratoCAB.IdContrato And ( Estado is null or Estado=1)), 0 ) as int) as TratosNoPagados ,cast(isnull((select  top 1  EPProyectado from INDEPAvanceCAB where IdContrato=dbo.INDEPContratoCAB.IdContrato and Nfactura is not null order by NumeroPago desc),0) as int) as EPProyectado ,(case dbo.INDEPContratoCAB.CodDocCobro  when 'FAVT' then 'FCEL' when 'FAEX' then 'C023' else 'FACC' end) as TipoDoc FROM dbo.INDEPContratoCAB  INNER JOIN dbo.INDTrContratistas ON dbo.INDEPContratoCAB.RutContratista = dbo.INDTrContratistas.RutContratista  INNER JOIN dbo.INDMaeUNegocioActivas ON dbo.INDEPContratoCAB.CtoEmpresa = dbo.INDMaeUNegocioActivas.CtoEmpresa AND dbo.INDEPContratoCAB.CtoCodigo = dbo.INDMaeUNegocioActivas.CtoCodigo  INNER JOIN dbo.maeEmpresa ON dbo.INDMaeUNegocioActivas.CtoEmpresa = dbo.maeEmpresa.empCodigo  WHERE ((dbo.INDEPContratoCAB.AnoControl = '2020')or  (dbo.INDEPContratoCAB.AnoControl = '2021')or  (dbo.INDEPContratoCAB.AnoControl = '2022') ) and (  cast(ROUND(ISNULL(dbo.INDEPContratoCAB.TotalContrato, 0) - ISNULL ((SELECT SUM(TotalTratos) AS Expr1   FROM dbo.INDEPAvanceCAB AS INDEPAvanceCAB_1  WHERE (IdContrato = dbo.INDEPContratoCAB.IdContrato) and estado = 5), 0),0) as int) )>0 ORDER BY dbo.INDEPContratoCAB.CtoEmpresa,dbo.INDEPContratoCAB.CtoCodigo,dbo.INDTrContratistas.RazonSocial")
Retencion=cursor.fetchall()
Retencion=pd.DataFrame(np.array(Retencion))

"""
Getting  Google Sheets
"""
CLIENT_SECRET_FILE = 'client_secret.json'
API_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

service = create_service(CLIENT_SECRET_FILE, API_NAME, API_VERSION,SCOPES )

google_sheets_id = '1uHsiEECIZlC6MYCrI3nPQCvqQMmGw_05O-79oPseBOo'


#inserta ComprasDescto
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
    range='Compras con Descto!A2:D'
    ).execute()
recordset = ComprasDescto.values.tolist()
"""
Insert rows
"""
request_body_values = construct_request_body(recordset)
service.spreadsheets().values().clear(spreadsheetId=google_sheets_id, range='Compras con Descto!A2:D').execute()
service.spreadsheets().values().update(
    spreadsheetId=google_sheets_id,
    valueInputOption='USER_ENTERED',
    range='Compras con Descto!A2:D',
    body=request_body_values
    ).execute()

print('ComprasDescto insertado completo')

#inserta Anticipo

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
    range='Anticipo de Contrato!A2:D'
    ).execute()
recordset = Anticipo.values.tolist()
"""
Insert rows
"""
request_body_values = construct_request_body(recordset)
service.spreadsheets().values().clear(spreadsheetId=google_sheets_id, range='Anticipo de Contrato!A2:D').execute()
service.spreadsheets().values().update(
    spreadsheetId=google_sheets_id,
    valueInputOption='USER_ENTERED',
    range='Anticipo de Contrato!A2:D',
    body=request_body_values
    ).execute()

print('Anticipo insertado completo')

#inserta Historico

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
    range='HISTORICO!A2:O'
    ).execute()
Historico[10]=Historico[10].astype(str)
recordset = Historico.values.tolist()
"""
Insert rows
"""
request_body_values = construct_request_body(recordset)
service.spreadsheets().values().clear(spreadsheetId=google_sheets_id, range='HISTORICO!A2:O').execute()
service.spreadsheets().values().update(
    spreadsheetId=google_sheets_id,
    valueInputOption='USER_ENTERED',
    range='HISTORICO!A2:O',
    body=request_body_values
    ).execute()

print('Historico insertado completo')

#inserta BDoriginal

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
    range='b2'
    ).execute()
recordset = BDoriginal.values.tolist()
"""
Insert rows
"""
request_body_values = construct_request_body(recordset)
service.spreadsheets().values().clear(spreadsheetId=google_sheets_id, range='b2').execute()
service.spreadsheets().values().update(
    spreadsheetId=google_sheets_id,
    valueInputOption='USER_ENTERED',
    range='b2',
    body=request_body_values
    ).execute()

print('BDoriginal insertado completo')

#inserta retenciones

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
    range='BD!A2:Q'
    ).execute()
recordset = Retencion.values.tolist()
"""
Insert rows
"""
request_body_values = construct_request_body(recordset)
service.spreadsheets().values().clear(spreadsheetId=google_sheets_id, range='BD!A2:Q').execute()
service.spreadsheets().values().update(
    spreadsheetId=google_sheets_id,
    valueInputOption='USER_ENTERED',
    range='BD!A2:Q',
    body=request_body_values
    ).execute()

print('Retenciones insertado completo')
cursor.close()