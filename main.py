"""
Ações para serem executadas no SAP, pode ter funções externas para pegar uma lista de Entradas a serem realizadas,
por exemplo.
"""
import sys

import ConectaSAP
from operator import itemgetter

import messagebox
from SAP import programarSQVI, retornarjobs
from janela import App
import auxiliares as aux


# Lista com as Views que deseja programar no SAP na transação SQVI e com a chave SQL equivalente
transacoes = [{'View': 'EKKN_COMP', 'Tipo': 'PEDIDOS'},
              {'View': 'EKET', 'Tipo': 'PEDIDOS'},
              {'View': 'EBAN', 'Tipo': 'REQUISIÇÕES DE COMPRA'}]
# Ordena a lista por Tipo para não ficar realizando SELECTs denecessários, pois deixa todos os tipos iguais "juntos"
transacoes = sorted(transacoes, key=itemgetter('Tipo'))

# Chamar o SAP com uma conexão padrão
ERP = ConectaSAP.RetornasessaoSAP('Teste')


# Opções de Execução de JOBs
mensagem = '1 - Gerar JOBs' + chr(13) + '2 - Retornar JOBs'
resposta = aux.criarinputbox('Opção de Ação:', mensagem, valorinicial='Seleciona a opção desejada')

app = App()

match resposta:
    case '1':
        # Gerar JOBs
        programarSQVI(transacoes, ERP.session, app)
    case '2':
        # Recuperar JOBs
        retornarjobs(transacoes, ERP.session, app)

    case _:
        messagebox.msgbox('Opção inválida! Favor verificar!', messagebox.MB_OK, 'Opção Inválida')
        sys.exit()
