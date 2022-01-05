import time
#pip install pypiwin32
#======================
import win32api
import win32com.client
#======================
import auxiliares as aux



class RetornasessaoSAP:
    def __init__(self, nomeexecutavel, conexaopadrao):
        self.fechasap = True
        self.fechaconexao = True
        self.fechasessao = True
        self.nomeexecutavel = nomeexecutavel
        self.conexaopadrao = conexaopadrao #'PRODUÇÃO ECC P03 - LOAD BALANCE'


    def definiraplicacao (self, name=self.nomeexecutavel):
        #Abre o SAP ou pega a instância do SAP já aberta
        self.fechasap = not(aux.process_exists(name))
        if not self.fechasap:
            win32api.ShellExecute(0, None, name, None, '', 1)
            time.sleep(3)

        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        if SapGuiAuto is not None:
            self.application = SapGuiAuto.GetScriptingEngine


    def definirconexao (self, name=self.nomeexecutavel, nameconnection=self.conexaopadrao):
        if self.application is None:
            #Chama a função que vai retornar a instância do SAP e a situação do mesmo (se estava aberto ou não)
            self.definiraplicacao(name)

        if self.application is not None:
            #Lista todas as conexões instaladas no SAP
            conexoes = aux.listaconexoes()
            if conexoes is not None:
                mensagem=''
                #Busca a conexão padrão na lista de conexões instaladas
                resposta = aux.pesquisalista(conexoes, nameconnection)
                #Teste para realizar ações caso não ache a conexão padrão na lista,
                #vai retornar uma lista de opções de conexões nesse caso
                if resposta == -1:
                    for indice, conexao in enumerate(conexoes):
                        mensagem = mensagem + str(indice + 1) + ' - ' + conexao + chr(13)
                    mensagem = mensagem + chr(13) * 2 + 'Escolha a conexão desejada:'
                    resposta = aux.criarinputbox('Escolha de Conexão', mensagem)
                    if str(resposta).isnumeric():
                        resposta = int(resposta) - 1

                #Verifica se a opção selecionada (ou que retornou da busca da conexão padrão) é uma opção válida
                if str(resposta).isnumeric() and resposta >= 1 and resposta <= len(conexoes):
                    #Finalmente pega a conexão informada (ou pela conexão padrão ou pela lista de seleção)
                    nomeconexao = conexoes(resposta)
                    #Verifica se já tem conexões ativas no SAP
                    if application.Connections.Count == 0 or self.fechasap:
                        #Caso a quantidade de conexões seja 0 ou o SAP estava fechado significa que é
                        #preciso abrir uma nova conexão
                        self.connection = self.application.OpenConnection(nomeconexao, True)
                        self.fechaconexao = True
                    else:
                        # Busca nas conexões ativas se são conexões do mesmo servidor dado como entrada
                        for conexao in application.Connections:
                            #Se encontra uma conexão do nome da conexão informado pára a busca
                            #(seja informado na conexão padrão ou na lista de opções)
                            if conexao.Description == nomeconexao:
                                self.connection = conexao
                                self.fechaconexao = False
                                break

                    #Caso as conexões existentes na aplicação não sejam da conexão informada
                    if self.connection is None:
                        self.connection = self.application.OpenConnection(nomeconexao, True)
                        self.fechaconexao = True


    def definirsessao(self, name=self.nomeexecutavel, nameconnection=self.conexaopadrao):
        if self.application is None:
            #Chama a função que vai retornar a instância do SAP e a situação do mesmo (se estava aberto ou não)
            self.definiraplicacao(name)

        if self.application is not None:
            #Lista todas as conexões instaladas no SAP
            conexoes = aux.listaconexoes()
            if self.connection is not None:








