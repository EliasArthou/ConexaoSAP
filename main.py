"""
Ações para serem executadas no SAP, pode ter funções externas para pegar uma lista de Entradas a serem realizadas,
por exemplo.
"""

import ConectaSAP
import auxiliares as aux
import sys
import messagebox
# from IPython.display import display

intervalo = aux.criarinputbox('Quantidade de Itens', 'Insira a quantidade de itens por job a ser extraído:')
if intervalo.isnumeric():
    intervalo = int(intervalo)
else:
    messagebox.msgbox('Valor inválido para a quantidade de itens!', messagebox.MB_OK, 'Erro quantidade de itens')
    sys.exit()

# Chamar o SAP com uma conexão padrão
SAP = ConectaSAP.RetornasessaoSAP('Teste')
# A sessão é carregada numa variável para facilitar para escrever o código,
# visto que o objeto será digitado várias vezes.
se = SAP.session

# Seleciona a transação desejada
se.findById("wnd[0]/tbar[0]/okcd").text = "SQVI"
# Confirma a transação
se.findById("wnd[0]").sendVKey(0)
# Seleciona a View
se.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "Eban"
# Confirma a seleção
se.findById("wnd[0]/usr/btnP1").press()

# Conexão com o Banco de Dados
conec = aux.Conec()

# Consulta a ser realizada no banco (será usado para fazer a seleção no SAP)
lista = conec.consulta("SELECT DISTINCT [Nº doc.ref] FROM [BDSIRI].[UsrBDSIRI].[GIG Analise Compromisso]"
                       " WHERE UPPER ([Ctg.val])='PEDIDOS' ORDER BY [Nº doc.ref] DESC")

# Quebra a lista em lista menores para o tamanho com a quantidade de item definido na variável intervalo
sublistas = list(aux.chunks(lista, intervalo))
# 'Looping' para quebrar o pedido em vários 'jobs'
for indice, item in enumerate(sublistas):
    # Carrega os n itens (definido na variável intervalo) para a memória para ser "colado" depois
    aux.list_to_clipboard(item, indice+1)
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

# Trata a finalização do SAP
if SAP.session is not None:
    SAP.finalizarsap()
