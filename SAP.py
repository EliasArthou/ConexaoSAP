"""
Rotinas do SAP serão executadas por aqui
"""
import messagebox
import sys
import datetime
from openpyxl import Workbook, load_workbook
import os
import auxiliares as aux
from janela import App


def programarSQVI(transacoes, se):
    """
    :param transacoes: lista contendo o dicionário com a transação e a classificação dos itens pesquisados.
    :param se: sessão do SAP que as ações do SAP será executado
    """
    # Variável para controlar se o Tipo mudou em relação à transação anterior para saber se será necessário um novo SELECT
    tipoatual = ''
    # Lista que receberá o resultado do SELECT
    lista = []
    # Inicia a variável
    datainicio = ''

    # Variável que receberá o intervalo de itens para a execução de cada JOB
    intervalo = aux.criarinputbox('Quantidade de Itens', 'Insira a quantidade de itens por job a ser extraído:',
                                  valorinicial='10000')
    # Verifica se foi digitado corretamente
    if intervalo.isnumeric():
        # Formata a entrada para número
        intervalo = int(intervalo)
    else:
        # Mensagem de erro e finalização do processo se o intervalo foi digitado errado
        messagebox.msgbox('Valor inválido para a quantidade de itens!', messagebox.MB_OK, 'Erro quantidade de itens')
        sys.exit()

    # ‘Looping’ para executar todas as visões solicitadas pelo usuário
    for transacao in transacoes:
        # Seleciona a transação desejada
        se.findById("wnd[0]/tbar[0]/okcd").text = "SQVI"
        # Confirma a transação
        se.findById("wnd[0]").sendVKey(0)
        # Seleciona a View
        se.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = transacao[list(transacao)[0]]
        # Confirma a seleção
        se.findById("wnd[0]/usr/btnP1").press()

        # Conexão com o Banco de Dados
        conec = aux.Conec()

        if transacao[list(transacao)[1]] != tipoatual:
            tipoatual = transacao[list(transacao)[1]]
            # Consulta a ser realizada no banco (será usado para fazer a seleção no SAP)
            lista = conec.consulta("SELECT DISTINCT [Nº doc.ref] FROM [BDSIRI].[UsrBDSIRI].[GIG Analise Compromisso]"
                                   " WHERE UPPER ([Ctg.val])='" + tipoatual + "' ORDER BY [Nº doc.ref]")

        if len(lista) == 0:
            messagebox.msgbox('Sem item para analisar com o tipo informado!', messagebox.MB_OK, 'Tabela Vazia')
        else:
            # Quebra a lista em lista menores para o tamanho com a quantidade de item definido na variável intervalo
            sublistas = list(aux.chunks(lista, intervalo))
            indice = 0
            # app = App(transacao[list(transacao)[0]], indice, len(sublistas))
            # app.mainloop()
            # 'Looping' para quebrar o pedido em vários 'jobs'
            for indice, item in enumerate(sublistas):
                # Grava a data e hora que começou a rodar os 'JOBs'
                datainicio = datetime.datetime.now()
                # Carrega os n itens (definido na variável intervalo) para a memória para ser "colado" depois
                aux.list_to_clipboard(item, indice + 1)
                if tipoatual == 'PEDIDOS':
                    # Para abrir a lista de colagem a caixa de texto não pode estar vazia (problema dessa transação específica)
                    se.findById("wnd[0]/usr/ctxtEBELN-LOW").text = "0"
                    # Abre a lista de colagem de itens da memória
                    se.findById("wnd[0]/usr/btn%_EBELN_%_APP_%-VALU_PUSH").press()
                else:
                    # Para abrir a lista de colagem a caixa de texto não pode estar vazia (problema dessa transação específica)
                    se.findById("wnd[0]/usr/ctxtBANFN-LOW").text = "0"
                    # Abre a lista de colagem de itens da memória
                    se.findById("wnd[0]/usr/btn%_BANFN_%_APP_%-VALU_PUSH").press()

                # Abre a janela para colar as informações da memória
                if se.ActiveWindow.name == "wnd[1]":
                    # Limpa a lista
                    se.findById("wnd[1]/tbar[0]/btn[16]").press()
                    # Carrega a lista com os itens da memória
                    se.findById("wnd[1]/tbar[0]/btn[24]").press()
                    # "Confirma" a lista
                    se.findById("wnd[1]/tbar[0]/btn[8]").press()
                # Selecionar o executar em 'background'
                if se.ActiveWindow.name == "wnd[0]":
                    # Aperta o botão de programar 'JOBs'
                    se.findById("wnd[0]/mbar/menu[0]/menu[2]").select()
                    if se.ActiveWindow.name == "wnd[1]":
                        # Confirma a "inpressora" utilizada (normalmente só virtual)
                        se.findById("wnd[1]/tbar[0]/btn[13]").press()
                        # Seleciona que é imediata
                        se.findById("wnd[1]/usr/btnSOFORT_PUSH").press()
                        # Salva o 'JOB'
                        se.findById("wnd[1]/tbar[0]/btn[11]").press()

            # Sai da transação
            se.EndTransaction()
            # Grava a data e hora que terminou de rodar os 'JOBs'
            datafim = datetime.datetime.now()
            # Verifica se o arquivo de LOG existe
            if not os.path.isfile('JobLog.xlsx'):
                wb = Workbook()
                wb.save('JobLog.xlsx')

            if os.path.isfile('JobLog.xlsx'):
                wb = load_workbook('JobLog.xlsx')

                ws = wb.active

                print([datainicio.strftime("%d/%m/%Y"), datainicio.strftime("%X"), datafim.strftime("%d/%m/%Y"),
                       datafim.strftime("%X"), se.info.user, transacao[list(transacao)[0]]])
                ws.append([datainicio.strftime("%d/%m/%Y"), datainicio.strftime("%X"), datafim.strftime("%d/%m/%Y"),
                           datafim.strftime("%X"), se.info.user, transacao[list(transacao)[0]]])

                wb.save('JobLog.xlsx')
