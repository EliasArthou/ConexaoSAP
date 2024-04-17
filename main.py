"""
Ações para serem executadas no SAP, pode ter funções externas para pegar uma lista de Entradas a serem realizadas,
por exemplo.
"""

import ConectaSAP
import messagebox as msg
from SAP import processosap
from janela import App
import auxiliares as aux
from gradus import Gradus
import pandas as pd
import time

testarorcamento = True
resultado = None
OIacriar = aux.retornarlista(10,  somentenaolidos=True)

if testarorcamento:
    api = Gradus('MW5S', 'P')
    # Tempo Inicial
    start_time = time.time()
    # Chama a API de Orçamento da GRADUS
    nome, resultado, _ = api.multifuncoes('Orçado Aprovado', limit=5000000, page=1, year=2024)
    # Tempo Final
    end_time = time.time()
    # Calcula diferença de tempo
    tempo_total = end_time - start_time

    # Converter para o formato MM:SS
    tempo_formatado = time.strftime("%M:%S", time.gmtime(tempo_total))
    # Mostra o tempo de extração do orçamento
    print("Tempo total para retornar o Orçamento:", tempo_formatado)

    resultado['Total'] = resultado['[values].[JANUARY]'] + resultado['[values].[FEBRUARY]'] + resultado['[values].[MARCH]'] + resultado['[values].[APRIL]'] + resultado['[values].[MAY]'] + resultado['[values].[JUNE]'] + resultado['[values].[JULY]'] + resultado['[values].[AUGUST]'] + resultado['[values].[SEPTEMBER]'] + resultado['[values].[OCTOBER]'] + resultado['[values].[NOVEMBER]'] + resultado['[values].[DECEMBER]']
    resultado.rename(columns={'accountingAccount': 'Conta Contábil'}, inplace=True)
    resultado = resultado[resultado['Conta Contábil'].str.contains('Investimento', case=False)]

    # Agrupar os dados por cidade e calcular a soma das idades
    somaporPRJ = resultado.groupby('Conta Contábil')['Total'].sum().reset_index()
    # Retira os espaços em branco
    somaporPRJ['Conta Contábil'] = somaporPRJ['Conta Contábil'].str.strip()
    # Cortar pra pegar só o PRJ
    somaporPRJ['Conta Contábil'] = somaporPRJ['Conta Contábil'].str[:11]
    # Criar um novo DataFrame com a coluna "Cidade" e a soma das idades
    resultado = pd.DataFrame({'Projeto': somaporPRJ['Conta Contábil'], 'Valor': somaporPRJ['Total']})

# Chamar o SAP com uma conexão padrão
if len(OIacriar) > 0:
    # retornoitem, respostaitem = aux.validadados(item)
    ERP = ConectaSAP.RetornasessaoSAP('PRD - ECC BR Produção')
    app = App()
    if len(ERP.session.info.user) > 0:
        processosap(ERP, app, OIacriar, resultado)
else:
    msg.msgbox('Sem e-mails de Ordem encontrado!', msg.MB_OK, 'Sem OIs')


