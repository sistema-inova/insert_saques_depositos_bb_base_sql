import time
import pyodbc
from classes.sisfin import Sisfin
from datetime import date, timedelta
from credenciais import *
import getpass
import os


def capturar_dados_sisfin_saques_depositos_bb_bacen():
    print("")
    escrever_cabecalho('CAPTURA DADOS SAQUES/DEPÓSITOS BB/BACEN')
    
    usuario = input('Matrícula: ')
    senha = getpass.getpass('Senha: ')

    for dia in range(-1, 11):
        if dia == -1:
            periodo = f"H-1"
            data_periodo = date.today() - timedelta(days=1)
        elif dia == 0:
            periodo = f"H"
            data_periodo = date.today()
        else:
            periodo = f"H+{dia}"
            data_periodo = date.today() + timedelta(days=dia)

    sisfin = Sisfin(usuario, senha)



    sisfin.sisfin_finalizar()




    # conn = pyodbc.connect(f"DRIVER={{SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USER};PWD={PASSWORD}")

    # # cursor = conn.cursor()

    # string = "CREATE TABLE TestTable(symbol varchar(15), leverage double, shares integer, price double)"
    # cur.execute(string)


    # for row in cursor:
    #     print(row)

def escrever_cabecalho(mensagem):
    print("{:*^100}" .format("*"))
    print("{:*^100}" .format(f" {mensagem} "))
    print("{:*^100}" .format("*"))
    print("")


if __name__ == "__main__":
    capturar_dados_sisfin_saques_depositos_bb_bacen()