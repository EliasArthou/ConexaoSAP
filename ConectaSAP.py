import time
# pip install pypiwin32
# ======================
import win32api
import win32com.client
# ======================
import auxiliares as aux
import messagebox


class RetornasessaoSAP:
    def __init__(self, conexaopadrao):
        self.application = None
        self.connection = None
        self.session = None
        self.fechasap = True
        self.fechaconexao = True
        self.fechasessao = True
        self.login = ''
        self.senha = ''
        self.nomeexecutavel = 'saplogon.exe'
        self.conexaopadrao = conexaopadrao  # 'PRODUÇÃO ECC P03 - LOAD BALANCE'

    def definiraplicacao(self):
        # Abre o SAP ou pega a instância do SAP já aberta
        self.fechasap = not (aux.process_exists(self.nomeexecutavel))
        if self.fechasap:
            win32api.ShellExecute(0, None, self.nomeexecutavel, None, '', 1)
            time.sleep(3)

        sapguiauto = win32com.client.GetObject('SAPGUI')
        if sapguiauto is not None:
            self.application = sapguiauto.GetScriptingEngine

    def definirconexao(self):
        if self.application is None:
            # Chama a função que vai retornar a instância do SAP e a situação do mesmo (se estava aberto ou não)
            self.definiraplicacao()

        if self.application is not None:
            # Lista todas as conexões instaladas no SAP
            if self.connection is None:
                conexoes = aux.listaconexoes()
                if conexoes is not None:
                    mensagem = ''
                    # Busca a conexão padrão na lista de conexões instaladas
                    resposta = aux.pesquisalista(conexoes, self.conexaopadrao)
                    # Teste para realizar ações caso não ache a conexão padrão na lista,
                    # vai retornar uma lista de opções de conexões nesse caso
                    if resposta == -1:
                        for indice, conexao in enumerate(conexoes):
                            mensagem = mensagem + str(indice + 1) + ' - ' + conexao + chr(13)
                        mensagem = mensagem + chr(13) * 2 + 'Escolha a conexão desejada:'
                        resposta = aux.criarinputbox('Escolha de Conexão', mensagem)
                        if str(resposta).isnumeric():
                            resposta = int(resposta) - 1

                    # Verifica se a opção selecionada (ou que retornou da busca da conexão padrão) é uma opção válida
                    if str(resposta).isnumeric() and 1 <= resposta <= len(conexoes):
                        # Finalmente pega a conexão informada (ou pela conexão padrão ou pela lista de seleção)
                        nomeconexao = conexoes[resposta]
                        # Verifica se já tem conexões ativas no SAP
                        if self.application.Connections.Count == 0 or self.fechasap:
                            # Caso a quantidade de conexões seja 0 ou o SAP estava fechado significa que é
                            # preciso abrir uma nova conexão
                            self.connection = self.application.OpenConnection(nomeconexao, True)
                            self.fechaconexao = True
                        else:
                            # Busca nas conexões ativas se são conexões do mesmo servidor dado como entrada
                            for conexao in self.application.Connections:
                                # Se encontra uma conexão do nome da conexão informado pára a busca
                                # (seja informado na conexão padrão ou na lista de opções)
                                if conexao.Description == nomeconexao:
                                    self.connection = conexao
                                    self.fechaconexao = False
                                    break

                        # Caso as conexões existentes na aplicação não sejam da conexão informada
                        if self.connection is None:
                            self.connection = self.application.OpenConnection(nomeconexao, True)
                            self.fechaconexao = True

    def definirsessao(self):
        if self.application is None:
            # Chama a função que vai retornar a instância do SAP e a situação do mesmo (se estava aberto ou não)
            self.definiraplicacao()

        if self.application is not None:
            # Verifica se tem conexão válida
            if self.connection is None:
                self.definirconexao()

            if self.connection is not None:
                # Verifica se o SAP não tenha nenhuma sessão ativa ou se a conexão foi criada no processo,
                # isso faria automaticamente não ter sessões já abertas pra conexão que foi dada como entrada.
                if self.connection.Sessions.Count == 0:
                    self.connection.Sessions(0).createSession
                    self.session = self.connection.Sessions(0)
                    self.fechasessao = True
                else:
                    for sessao in self.connection.sessions:
                        # Verifica se a sessão está na tela de login
                        if sessao.info.transaction == 'S000' or sessao.info.transaction == "SESSION_MANAGER":
                            self.session = sessao
                            self.fechasessao = False
                            break

                    if self.session is None:
                        if self.connection.Sessions.Count < 4:
                            self.connection.Sessions(self.connection.Sessions.Count - 1).createSession
                            self.session = self.connection.Sessions(self.connection.Sessions.Count - 1)
                        else:
                            messagebox.msgbox('Limite de Janelas atingido! Feche alguma janela para continuar!',
                                              messagebox.MB_OK, 'Limite de Janelas')

