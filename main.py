import pyautogui
import time
import pyodbc
from classes.sisfin import Sisfin
from credenciais import *
import getpass
import os

conn = pyodbc.connect(f"DRIVER={{SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USER};PWD={PASSWORD}")

cursor = conn.cursor()

# comando_insert = """
#     INSERT INTO [dbo].[TBL_EMPREGADOS]
#         ([matricula]
#         ,[nomeCompleto]
#         ,[primeiroNome]
#         ,[dataNascimento]
#         ,[codigoFuncao]
#         ,[nomeFuncao]
#         ,[codigoLotacaoAdministrativa]
#         ,[nomeLotacaoAdministrativa]
#         ,[codigoLotacaoFisica]
#         ,[nomeLotacaoFisica]
#         ,[created_at]
#         ,[updated_at])
#     VALUES
#         ('c098453'
#         ,'RAFAEL PIMENTEL GONCALVES'
#         ,'RAFAEL'
#         ,'1984-12-09'
#         ,null
#         ,null
#         ,7257
#         ,'GI ALIENAR BENS MOVEIS IMOV SAO PAULO,SP'
#         ,null
#         ,null
#         ,'2020-01-14 11:10:17.820'
#         ,'2020-06-03 16:19:05.000')
#     """
# cursor.execute(comando_insert)
# cursor.commit()

cursor.execute(f'SELECT * FROM {DATABASE}.[dbo].[TBL_EMPREGADOS]')

for row in cursor:
    print(row)