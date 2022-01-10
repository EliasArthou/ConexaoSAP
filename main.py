"""
Ações para serem executadas no SAP, pode ter funções externas para pegar uma lista de Entradas a serem realizadas,
por exemplo.
"""
import ConectaSAP
import auxiliares as aux

'''
conec = aux.Conec()
teste = conec.consulta("SELECT * FROM [GIG Arvores Conta]", True)
print(teste)

'''

SAP = ConectaSAP.RetornasessaoSAP('Teste')

if SAP.session is not None:
    SAP.finalizarsap()
