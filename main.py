import ConectaSAP
import auxiliares as aux

SAP = ConectaSAP.RetornasessaoSAP('PRODUÇÃO ECC P03 - LOAD BALANCE')
SAP.definirsessao()

#if SAP.session.info.transaction == 'S000':
SAP.login = aux.criarinputbox('Usuário', 'Digite o Login:')
SAP.senha = aux.criarinputbox('Senha', 'Digite o Senha:', '*')
if len(SAP.login) > 0:
    SAP.session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = SAP.login
if len(sessao.senha) > 0:
    SAP.session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = SAP.senha

SAP.session.findById("wnd[0]").sendVKey(0)
