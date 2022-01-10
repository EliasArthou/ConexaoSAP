
import time
# pip install pypiwin32
# ======================
import win32api
import win32com.client
# ======================
import auxiliares as aux
import messagebox


class RetornasessaoSAP:
    """
    Criar objeto de conexão com o SAP
    """
    def __init__(self, conexaopadrao, nomeexecutavel='saplogon.exe'):
        self.application = None
        self.connection = None
        self.session = None
        self.fechasap = True
        self.fechaconexao = True
        self.fechasessao = True
        self.login = ''
        self.senha = ''
        self.nomeexecutavel = nomeexecutavel # 'saplogon.exe'
        self.conexaopadrao = conexaopadrao  # 'PRODUÇÃO ECC P03 - LOAD BALANCE'
        self.definirsessao()

    def definiraplicacao(self):
        """
         Abre o SAP ou pega a instância do SAP já aberta
        """
        self.fechasap = not (aux.process_exists(self.nomeexecutavel))
        if self.fechasap:
            win32api.ShellExecute(0, None, self.nomeexecutavel, None, '', 1)
            # Ponto de atenção, espera o executável do SAP carregar, pode variar de máquina para máquina
            time.sleep(3)

        sapguiauto = win32com.client.GetObject('SAPGUI')
        if sapguiauto is not None:
            self.application = sapguiauto.GetScriptingEngine

    def definirconexao(self):
        """
        Define a conexão do SAP que será utilizada para executar a ação desejada
        """
        if self.application is None:
            # Chama a função que vai retornar a instância do SAP e a situação do mesmo (se estava aberto ou não)
            self.definiraplicacao()

        # Verifica se a aplicação (SAP) foi alocada na memória
        if self.application is not None:
            if self.connection is None:
                # Lista todas as conexões instaladas no SAP
                conexoes = aux.listaconexoes()
                if conexoes is not None:
                    # Verifica se já tem conexões ativas no SAP
                    if self.application.Connections.Count > 0:
                        nomeconexao = self.conexaopadrao
                        # Busca nas conexões ativas se são conexões do mesmo servidor dado como entrada
                        for conexao in self.application.Connections:
                            # Se encontra uma conexão com o nome informado para a busca
                            # (seja informado no nome de entrada ou na lista de opções)
                            if conexao.Description == nomeconexao:
                                self.connection = conexao
                                self.fechaconexao = False
                                break

                        nomeconexao = ''

                    if self.connection is None:
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

                        # Verifica se a opção selecionada (ou que retornou da busca da conexão padrão)
                        # é uma opção válida
                        if str(resposta).isnumeric() and 1 <= resposta <= len(conexoes):
                            # Finalmente pega a conexão informada (pela conexão padrão ou pela lista de seleção)
                            nomeconexao = conexoes[resposta]
                            # Verifica se já tem conexões ativas no SAP
                            if self.application.Connections.Count == 0 or self.fechasap:
                                # Caso a quantidade de conexões seja 0 ou o SAP estava fechado significa ser
                                # preciso abrir uma nova conexão
                                self.connection = self.application.OpenConnection(nomeconexao, True)
                                self.fechaconexao = True
                            else:
                                # Busca nas conexões ativas se são conexões do mesmo servidor dado como entrada
                                for conexao in self.application.Connections:
                                    # Se encontra uma conexão com o informado para a busca
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
        """
        Define a sessão na conexão informada, caso já exista utiliza a mesma, caso não cria a sessão.
        """
        if self.application is None:
            # Chama a função que vai retornar a instância do SAP e a situação do mesmo (se estava aberto ou não)
            self.definiraplicacao()

        if self.application is not None:
            # Verifica se tem conexão válida
            if self.connection is None:
                self.definirconexao()

            if self.connection is not None:
                # Verifica se o SAP não tenha nenhuma sessão ativa ou se a conexão foi criada no processo,
                # isso faria automaticamente não ter sessões já abertas para conexão dada como entrada.
                if self.connection.Sessions.Count == 0:
                    self.connection.sessions(0).createSession()
                    self.session = self.connection.Sessions(0)
                    self.fechasessao = True
                else:
                    for sessao in self.connection.Sessions:
                        # Verifica se a sessão está na tela de "login"
                        if sessao.info.transaction == 'S000' or sessao.info.transaction == "SMEN" \
                                or sessao.info.transaction == "SESSION_MANAGER":
                            self.session = sessao
                            self.fechasessao = False
                            break

                    if self.session is None:
                        if self.connection.Sessions.Count < 4:
                            sessaotemp = self.connection.sessions(self.connection.Sessions.Count-1)
                            sessaotemp.createSession()
                            time.sleep(1)
                            self.session = self.connection.sessions(self.connection.Sessions.Count-1)
                            self.fechasessao = True
                        else:
                            messagebox.msgbox('Limite de Janelas atingido! Feche alguma janela para continuar!',
                                              messagebox.MB_OK, 'Limite de Janelas')

                    else:
                        if self.session.info.transaction == 'S000':
                            self.login = aux.criarinputbox('Usuário', 'Digite o Login:')
                            self.senha = aux.criarinputbox('Senha', 'Digite o Senha:', '*')
                            if len(self.login) > 0:
                                self.session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = self.login
                            if len(self.senha) > 0:
                                self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = self.senha
                            self.session.findById("wnd[0]").sendVKey(0)

    def finalizarsap(self):
        """
        Encerra o SAP conforme iniciou o processo, se iniciou fechado, fecha o SAP. Se tinha a conexão aberta,
        a mantém, etc.
        """
        if self.fechasap:
            aux.fecharprograma('saplogon.exe')
        elif self.fechaconexao:
            self.session.findById("wnd[0]").close()
        elif self.session:
            self.session.findById("wnd[0]").close()
