import ConectaSAP
import auxiliares as aux

conec = aux.Conec()
teste = conec.consulta("SELECT * FROM [GIG Arvores Conta]", True)
print(teste)

'''
SAP = ConectaSAP.RetornasessaoSAP('Teste')

if SAP.session is not None:


if SAP.fechasap:
    aux.fecharprograma('saplogon.exe')
elif SAP.fechaconexao:
    SAP.session.findById("wnd[0]").close()
elif SAP.session:
    SAP.session.findById("wnd[0]").close()

'''