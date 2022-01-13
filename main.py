"""
Ações para serem executadas no SAP, pode ter funções externas para pegar uma lista de Entradas a serem realizadas,
por exemplo.
"""

import ConectaSAP
from operator import itemgetter
import SAP


# Lista com as Views que deseja programar no SAP na transação SQVI e com a chave SQL equivalente
transacoes = [{'View': 'EKKN_COMP', 'Tipo': 'PEDIDOS'},
              {'View': 'EKET', 'Tipo': 'PEDIDOS'},
              {'View': 'EBAN', 'Tipo': 'REQUISIÇÕES DE COMPRA'}]
# Ordena a lista por Tipo para não ficar realziando SELECTs denecessários, pois deixa todos os tipos iguais "juntos"
transacoes = sorted(transacoes, key=itemgetter('Tipo'))

# Chamar o SAP com uma conexão padrão
ERP = ConectaSAP.RetornasessaoSAP('Teste')

# A sessão é carregada numa variável para facilitar para escrever o código,
# visto que o objeto será digitado várias vezes.
se = ERP.session

SAP.programarSQVI(transacoes, se)

# Trata a finalização do SAP
if se is not None:
    ERP.finalizarsap()





