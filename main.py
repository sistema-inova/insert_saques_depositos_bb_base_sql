import time
import pyodbc
from classes.sisfin import Sisfin
from datetime import date, timedelta, datetime
from credenciais import *
import getpass
import os


def capturar_dados_sisfin_saques_depositos_bb_bacen():
    print("")
    escrever_cabecalho('CAPTURA DADOS SAQUES/DEPÓSITOS BB/BACEN')
    
    usuario = input('Matrícula: ')
    senha = getpass.getpass('Senha: ')
    data_hora_consulta = str(f"{datetime.now():%Y-%m-%d %H:%M:%S}")
    sisfin = Sisfin(usuario, senha)

    conexao_banco = pyodbc.connect(f"DRIVER={{SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USER};PWD={PASSWORD}")
    cursor = conexao_banco.cursor()

    if not cursor.tables(table='TP_90003_03_TEMP_SAQUE_DEPOSITO_BB_BACEN', tableType='TABLE').fetchone():
        consulta = "CREATE TABLE [dbo].[TP_90003_03_TEMP_SAQUE_DEPOSITO_BB_BACEN]([dtDataMovimento] [date], [vcTipoMovimentacao] [varchar](10), [chLinhaVmcbSisfin] [char](3), [inCodigoUnidade] [int], [vcNomeUnidade] [varchar](max), [chLM] [char](2), [dcValorSaque] [decimal](17,2), [vcSL] [varchar](10), [dtDataCaptura] [datetime])"
        cursor.execute(consulta)
        conexao_banco.commit()

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

        print("")
        escrever_cabecalho('SAQUE BB/BACEN')
        sisfin.vmcb_acessar(periodo, 'S')
        tela = 1
        terminou_agencias = False
        tipo_movimentacao = "SAQUE"
        # TRATAMENTO CASO NÃO EXISTA MOVIMENTAÇÃO NO DIA
        if  sisfin.sisfin_ler(5, 12, 4).strip() != 'Saque':
            sisfin.vmcb_tratamento_sem_movimentacao()
        else:
            while not terminou_agencias:
                linha = 9
                for passo in range(1, linha + 2):
                    time.sleep(0.5)
                    if sisfin.sisfin_ler(linha, 6, 4).strip() != '':
                        codigo_vmcb_sisfin = sisfin.sisfin_ler(linha, 1, 2).strip()
                        codigo_unidade = sisfin.sisfin_ler(linha, 6, 3).strip()
                        nome_unidade = sisfin.sisfin_ler(linha, 13, 32).strip()
                        lm = sisfin.sisfin_ler(linha, 47, 1).strip()
                        valor_saque = sisfin.sisfin_ler(linha, 50, 17).strip().replace(".", "").replace(",", ".")
                        sl = sisfin.sisfin_ler(linha, 69, 3).strip()
                        realiza_insert_banco(conexao_banco, cursor, data_periodo, tipo_movimentacao, codigo_vmcb_sisfin, codigo_unidade, nome_unidade, lm, valor_saque, sl, data_hora_consulta)
                        print(f"Linha: {codigo_vmcb_sisfin} | Código Unidade: {codigo_unidade} | Nome Unidade: {nome_unidade} | LM: {lm} | Valor do Saque: R$ {valor_saque} | SL: {sl}")
                    else:
                        terminou_agencias = True
                        break
                    linha += 1
                sisfin.enter()
                tela += 1
            sisfin.sisfin_voltar_pagina_inicial_cardapio(20)

        print("")
        escrever_cabecalho('DEPOSITO BB/BACEN')
        sisfin.vmcb_acessar(periodo, 'D')
        # TRATAMENTO CASO NÃO EXISTA MOVIMENTAÇÃO NO DIA
        sem_movimentacao = sisfin.sisfin_ler(5, 12, 7).strip() != 'Deposito'
        if sem_movimentacao:
            sisfin.vmcb_tratamento_sem_movimentacao()
        else:
            terminou_agencias = False
            tipo_movimentacao = "DEPOSITO"
            while not terminou_agencias:
                linha = 9
                for passo in range(1, linha + 2):
                    time.sleep(0.5)
                    if sisfin.sisfin_ler(linha, 6, 3).strip() != '':
                        codigo_vmcb_sisfin = sisfin.sisfin_ler(linha, 1, 2).strip()
                        codigo_unidade = sisfin.sisfin_ler(linha, 6, 3).strip()
                        nome_unidade = sisfin.sisfin_ler(linha, 13, 32).strip()
                        lm = sisfin.sisfin_ler(linha, 47, 1).strip()
                        valor_saque = sisfin.sisfin_ler(linha, 50, 17).strip().replace(".", "").replace(",", ".")
                        sl = sisfin.sisfin_ler(linha, 69, 3).strip()
                        realiza_insert_banco(conexao_banco, cursor, data_periodo, tipo_movimentacao, codigo_vmcb_sisfin, codigo_unidade, nome_unidade, lm, valor_saque, sl, data_hora_consulta)
                        print(f"Linha: {codigo_vmcb_sisfin} | Código Unidade: {codigo_unidade} | Nome Unidade: {nome_unidade} | LM: {lm} | Valor do Depósito: R$ {valor_saque} | SL: {sl}")
                    else:
                        terminou_agencias = True
                        break
                    linha += 1
            sisfin.enter()
            sisfin.sisfin_voltar_pagina_inicial_cardapio(20)

    sisfin.sisfin_finalizar()


def escrever_cabecalho(mensagem):
    print("{:*^120}" .format("*"))
    print("{:*^120}" .format(f" {mensagem} "))
    print("{:*^120}" .format("*"))
    print("")


def realiza_insert_banco(conexao_banco, cursor, data_periodo, tipo_movimentacao, codigo_vmcb_sisfin, codigo_unidade, nome_unidade, lm, valor_saque, sl, data_hora_consulta):
    consulta = """
        INSERT INTO [dbo].[TP_90003_03_TEMP_SAQUE_DEPOSITO_BB_BACEN]
           ([dtDataMovimento]
           ,[vcTipoMovimentacao]
           ,[chLinhaVmcbSisfin]
           ,[inCodigoUnidade]
           ,[vcNomeUnidade]
           ,[chLM]
           ,[dcValorSaque]
           ,[vcSL]
           ,[dtDataCaptura])
        VALUES
            ('{}'
            ,'{}'
            ,'{}'
            ,'{}'
            ,'{}'
            ,'{}'
            ,'{}'
            ,'{}'
            ,'{}')
        """
    cursor.execute(consulta.format(data_periodo, tipo_movimentacao, codigo_vmcb_sisfin, codigo_unidade, nome_unidade, lm, valor_saque, sl, data_hora_consulta))
    conexao_banco.commit()


if __name__ == "__main__":
    capturar_dados_sisfin_saques_depositos_bb_bacen()