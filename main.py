import ConectaSAP

SAP = ConectaSAP.RetornasessaoSAP('Teste')#PRODUÇÃO ECC P03 - LOAD BALANCE


SAP.session.endtransaction()
