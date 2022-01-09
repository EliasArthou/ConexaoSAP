import ConectaSAP
import auxiliares as aux

SAP = ConectaSAP.RetornasessaoSAP('Teste')

if SAP.session is not None:
    SAP.session.endtransaction()

if SAP.fechasap:
    aux.fecharprograma('saplogon.exe')
elif SAP.fechaconexao:
    SAP.session.findById("wnd[0]").close()
elif SAP.session:
    SAP.session.findById("wnd[0]").close()

