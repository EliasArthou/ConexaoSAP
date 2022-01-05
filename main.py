import ConectaSAP as SAP

'''
SapGuiAuto = aux.abrirarquivo('saplogon.exe')
application = SapGuiAuto.GetScriptingEngine
print(application.connections.count)
connection = application.OpenConnection('PRODUÇÃO ECC P03 - LOAD BALANCE', True)
time.sleep(3)
session = connection.Children(0)
session.findByid('wnd[0]').maximize
session.findByid('wnd[0]/usr/txtRSYST-BNAME').text = 'oi234957'
session.findByid('wnd[0]/usr/pwdRSYST-BCODE').text = 'Efarthou04062018*EFA'
session.findByid('wnd[0]').sendVKey(0)

result = aux.process_exists('saplogon.exe')
print(result)
'''
SAP.retornasessaovalida()
