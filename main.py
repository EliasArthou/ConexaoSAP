import ConectaSAP
import auxiliares as aux

sessao = ConectaSAP.RetornasessaoSAP('PRODUÇÃO ECC P03 - LOAD BALANCE')
sessao.definirsessao()

if sessao.session.info.transaction == 'S000':
    sessao.login = aux.criarinputbox('Usuário', 'Digite o Login:')
    sessao.senha = aux.criarinputbox('Senha', 'Digite o Senha:', '*')
    if len(sessao.login) > 0:
        sessao.session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = sessao.login
    if len(sessao.senha) > 0:
        sessao.session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = sessao.senha

    sessao.session.findById("wnd[0]").sendVKey(0)
