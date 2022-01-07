import ConectaSAP

SAP = ConectaSAP.RetornasessaoSAP('Teste')

if SAP.session is not None:
    SAP.session.endtransaction()
