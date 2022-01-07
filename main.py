import ConectaSAP

SAP = ConectaSAP.RetornasessaoSAP('Teste', 'saplogon.exe')

if SAP.session is not None:
    SAP.session.endtransaction()
