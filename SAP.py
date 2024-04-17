"""
Rotinas do SAP serão executadas por aqui
"""

import datetime
import os.path
import numpy as np
import auxiliares as aux
import pandas as pd
import banco as bd
import getpass
from datetime import datetime, timedelta
import setupextrator as variaveis


# from janela import App

def processosap(se, visual, lista, listaorcamento=None):
    """
    :param visual: tela de retorno do usuário.
    :param se: sessão do SAP que as ações do SAP será executado
    :param lista: lista de e-mails
    :param listaorcamento: lista dos PRJs
    """
    sessao = None
    item = None

    # try:
    problemasap = False
    sessao = se.session
    banco = bd.Banco()
    visual.mudartexto('labeljob', 'Consultando Atualização', 'Arial 15 bold')
    df = banco.consultar(f"SELECT * FROM TblAtualizacaoes WHERE Relatorio = 'Ordem Interna'")

    # Verifique se o DataFrame está vazio
    if len(df) == 0:
        # Obtenha a data e hora atuais
        data_atual = datetime.now()

        # Subtrair um dia
        um_dia = timedelta(days=1)
        data_anterior = data_atual - um_dia

        # Calcule a diferença entre as datas
        diferenca = data_anterior - data_atual

    else:
        # Converta a série 'Atualizacao' para datetime, se necessário
        if not pd.api.types.is_datetime64_any_dtype(df['Atualizacao']):
            df['Atualizacao'] = pd.to_datetime(df['Atualizacao'])

        # Calcule a diferença entre a data da atualização e a data atual
        diferenca = datetime.now() - df['Atualizacao']

    # Extraia apenas o número de dias da diferença
    diferenca_em_dias = diferenca.dt.days

    # Extraia apenas o número de dias da diferença
    diferenca_em_dias = diferenca_em_dias[0]

    sessao.FindById("wnd[0]").maximize()
    # ================================= Extração de Ordens Internas =======================================
    if diferenca_em_dias >= 1:
        # Preenche a transação a KOK5
        sessao.FindById("wnd[0]/tbar[0]/okcd").Text = "kok5"
        # Inicia a transação
        sessao.FindById("wnd[0]").sendVKey(0)
        # Verifica se abriu a lista com variantes e fecha
        if sessao.Children.count > 1:
            sessao.findById("wnd[1]").sendVKey(12)
        # Define a variant que será utilizada
        sessao.findById("wnd[0]/usr/subSELEKTION:SAPLKOSM:0510/ctxtCODIA-VARIANT").text = "listaordens"
        # Executa a extração
        sessao.findById("wnd[0]").sendVKey(8)
        # Chama a variante
        sessao.findById("wnd[0]").sendVKey(33)
        # Pega a tela da lista de variante e armazena em uma variável
        tabela = sessao.findById(
            "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")
        # Inicia a lista do início
        tabela.firstVisibleRow = 0
        # Número total de linhas na tabela
        total_rows = tabela.RowCount
        # Nome da variante que você está procurando
        texto_procurado = "/FLA"

        # Loop pelas linhas da tabela
        for row_index in range(total_rows):
            # Obtém o texto da primeira coluna da linha atual
            texto_coluna1 = tabela.GetCellValue(row_index, "VARIANT")  # Supondo que a primeira coluna tenha o índice 1

            # Verifica se o texto da primeira coluna é o texto que você está procurando
            if texto_procurado in texto_coluna1:
                # Se encontrado, coloca no início da lista pela barra de rolagem
                tabela.firstVisibleRow = row_index
                # Seleciona a linha da variante encontrada
                tabela.currentCellRow = row_index
                # Clica no item para a seleção e execução
                tabela.clickCurrentCell()
                # Chama a exportação do arquivo
                sessao.findById("wnd[0]").sendVKey(9)
                # Seleciona como tipo "TXT"
                sessao.findById("wnd[1]").sendVKey(0)
                # Define o nome do arquivo
                sessao.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "listaordens.txt"
                # Confirma
                sessao.findById("wnd[1]").sendVKey(11)
                # Sai da transação
                sessao.EndTransaction()
                if os.path.isfile(os.path.join(aux.caminhospadroes(5), 'SAP/SAP GUI/', 'listaordens.txt')):
                    usuario = getpass.getuser()
                    banco.executarSQL('TRUNCATE TABLE [Lista Ordem Internas]')
                    retorno = aux.tratararquivo(os.path.join(aux.caminhospadroes(5), 'SAP/SAP GUI/', 'listaordens.txt'),
                                                retornadf=True)
                    # Removendo a coluna de tipo (ZD12)
                    retorno = retorno.drop(columns=['Tp.'])
                    # Define os nomes das colunas
                    novos_nomes = ['Ordem', 'Descrição da atividade', 'Centro de Lucro', 'Centro de Custo Responsável',
                                   'Centro de Custo Solicitante', 'Área de Aplicação', 'Objetivo Setorial',
                                   'Grupo de Ordem', 'Código do Cliente', 'Data Entrada', 'Status']

                    # Renomeia as colunas
                    retorno.columns = novos_nomes

                    # Adiciona o dataframe
                    banco.adicionardf('Lista Ordem Internas', retorno)
                    if len(df) == 0:
                        banco.executarSQL(
                            f"INSERT INTO TblAtualizacaoes VALUES ('Ordem Interna', '{datetime.now().strftime("%m/%d/%Y %H:%M:%S")}','{usuario}')")
                    else:
                        banco.executarSQL(
                            f"UPDATE TblAtualizacaoes SET Atualizacao = '{datetime.now().strftime("%m/%d/%Y %H:%M:%S")}', Usuario = '{usuario}' WHERE Relatorio = 'Ordem Interna'")

                break

            # Se o loop terminar sem encontrar o texto, imprima uma mensagem
            print("Texto não encontrado na tabela.")

    visual.acertaconfjanela(True)

    # ==================== Parte Gráfica =======================================================
    # Coloca o passo que se encontra
    visual.mudartexto('labeljob', 'Criando Ordem')
    # ==================== Parte Gráfica =======================================================
    for indice, item in enumerate(lista):
        problemasap = False
        retornoitem, respostaitem = aux.validadados(item)
        if not retornoitem:
            visual.mudartexto('labelpassos', f'Criando OI {indice + 1} de {len(lista)}...')
            visual.configurarbarra('barrajob', len(lista), indice + 1)
            if listaorcamento is not None:
                valororcamento = listaorcamento.loc[listaorcamento['Projeto'].str.strip() == item['Objetivo Setorial'].strip(), 'Valor'].iloc[0]
                if valororcamento.empty:
                   valororcamento = 0
            else:
                valororcamento = 1

            if valororcamento > 0:
                sessao.FindById("wnd[0]").maximize()
                sessao.FindById("wnd[0]/tbar[0]/okcd").Text = "ko01"
                sessao.FindById("wnd[0]").sendVKey(0)
                sessao.FindById("wnd[0]").restore()
                sessao.FindById("wnd[0]/usr/ctxtCOAS-AUART").Text = "zd12"
                sessao.FindById("wnd[0]").sendVKey(0)
                sessao.FindById("wnd[0]/usr/txtCOAS-KTEXT").Text = item["Descrição da atividade"].strip()[:40]
                sessao.FindById(
                    "wnd[0]/usr/tabsTABSTRIP_600/tabpBUT1/ssubAREA_FOR_601:SAPMKAUF:0601/subAREA1:SAPMKAUF:0315/ctxtCOAS-PRCTR").Text = \
                    item["Centro de Lucro"]
                print(len(sessao.FindById("wnd[0]/sbar").Text))
                sessao.FindById(
                    "wnd[0]/usr/tabsTABSTRIP_600/tabpBUT1/ssubAREA_FOR_601:SAPMKAUF:0601/subAREA1:SAPMKAUF:0315/ctxtCOAS-KOSTV").Text = \
                    item["Centro de Custos Responsável"]
                sessao.FindById(
                    "wnd[0]/usr/tabsTABSTRIP_600/tabpBUT1/ssubAREA_FOR_601:SAPMKAUF:0601/subAREA1:SAPMKAUF:0315/ctxtCOAS-AKSTL").Text = \
                    item["Centro de Custos Solicitante"]
                sessao.FindById("wnd[0]").maximize()
                sessao.FindById("wnd[0]/usr/tabsTABSTRIP_600/tabpBUT5").Select()
                item["Resposta"] = sessao.FindById("wnd[0]/sbar").Text
                if len(item['Resposta'].strip()) > 0:
                    visual.mudartexto('labelpassos', 'Enviando E-mail')
                    aux.reply_or_forward_email(item['Email'], item['Resposta'], variaveis.emails)
                    sessao.EndTransaction()
                else:
                    sessao.FindById(
                        "wnd[0]/usr/tabsTABSTRIP_600/tabpBUT5/ssubAREA_FOR_601:SAPMKAUF:0601/subAREA2:SAPLXAUF:0100/ctxtGLOBAL_AUFK-YYAREA_AP").Text = \
                        item["Área de Aplicação"]
                    sessao.FindById(
                        "wnd[0]/usr/tabsTABSTRIP_600/tabpBUT5/ssubAREA_FOR_601:SAPMKAUF:0601/subAREA2:SAPLXAUF:0100/ctxtGLOBAL_AUFK-YYGRP_ORD").Text = \
                        item["Grupo de Ordem"]
                    sessao.FindById("wnd[0]").maximize()
                    sessao.FindById("wnd[0]").restore()
                    sessao.FindById(
                        "wnd[0]/usr/tabsTABSTRIP_600/tabpBUT5/ssubAREA_FOR_601:SAPMKAUF:0601/subAREA2:SAPLXAUF:0100/ctxtGLOBAL_AUFK-YYAREA_AP").SetFocus()
                    sessao.FindById(
                        "wnd[0]/usr/tabsTABSTRIP_600/tabpBUT5/ssubAREA_FOR_601:SAPMKAUF:0601/subAREA2:SAPLXAUF:0100/ctxtGLOBAL_AUFK-YYAREA_AP").caretPosition = 4
                    sessao.FindById("wnd[0]").sendVKey(0)
                    sessao.FindById(
                        "wnd[0]/usr/tabsTABSTRIP_600/tabpBUT5/ssubAREA_FOR_601:SAPMKAUF:0601/subAREA2:SAPLXAUF:0100/ctxtGLOBAL_AUFK-YYOBJ_SET").Text = \
                        item["Objetivo Setorial"]
                    sessao.FindById(
                        "wnd[0]/usr/tabsTABSTRIP_600/tabpBUT5/ssubAREA_FOR_601:SAPMKAUF:0601/subAREA2:SAPLXAUF:0100/ctxtGLOBAL_AUFK-YYKUNNR").Text = \
                        item["Código do Cliente"]
                    sessao.FindById("wnd[0]/tbar[0]/btn[11]").press()
                    sessao.FindById("wnd[0]").maximize()
                    item["Resposta"] = sessao.FindById("wnd[0]/sbar").Text
                    if len(item['Resposta'].strip()) > 0:
                        visual.mudartexto('labelpassos', 'Enviando E-mail')
                        aux.reply_or_forward_email(item['Email'], item['Resposta'], variaveis.emails)
        else:
            visual.mudartexto('labelpassos', 'Enviando E-mail')
            aux.reply_or_forward_email(item['Email'], respostaitem, variaveis.emails, respostaitem)
        if sessao.info.transaction != "SESSION_MANAGER":
            sessao.EndTransaction()

    lista = pd.DataFrame(lista)
    lista['Ordem'] = lista['Resposta'].str.extract(r'(\d{10})')
    lista['Data Entrada'] = datetime.datetime.now().strftime('%d/%m/%Y')
    lista = lista.loc[lista['Resposta'].str.len() > 0]
    visual.mudartexto('labelpassos', 'Criação OIs Concluídas')
    # Retorna o que está fazendo em 'background'
    if sessao.info.transaction != "SESSION_MANAGER":
        # Sai da transação
        sessao.EndTransaction()
    # Grava a data e hora que terminou de rodar os 'JOBs'
    datafim = datetime.datetime.now()

    # Verifica se o arquivo de LOG existe para não ter erro quando abrir o arquivo para salvar o LOG
    # if os.path.isfile(pwd.arquivoOIs):

    colunaapagar = ['Resposta']
    listacampos = variaveis.camposdados
    listacampos.append(colunaapagar)
    # listacampos = ['Ordem',
    #                'Descrição da atividade',
    #                'Centro de Lucro',
    #                'Centro de Custo Responsável',
    #                'Centro de Custo Solicitante',
    #                'Área de Aplicação',
    #                'Objetivo Setorial',
    #                'Grupo de Ordem',
    #                'Código do Cliente',
    #                'Data Entrada',
    #                'Resposta']
    lista = lista[variaveis.camposdados]
    lista = lista.dropna(subset=['Ordem'])
    lista.replace('', np.nan, inplace=True)
    banco.adicionardf('Lista Ordem Internas', lista, colunaapagar)
    # except Exception as e:
    #     df = pd.DataFrame(lista)
    #     arquivodeerro = 'Criados.xlsx'
    #     if os.path.isfile(arquivodeerro):
    #         os.remove(arquivodeerro)
    #     df.to_excel(arquivodeerro, index=False)
