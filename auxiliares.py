"""
Funções complementares de ações gerais para auxiliar o funcionamento do processo.
"""
import sensiveis as pwd
import win32com.client
import re
import pandas as pd
import os
import banco as bd
import messagebox as msg
import subprocess
import time
from fuzzywuzzy import fuzz
import setupextrator as variaveis

# Código ANSI para iniciar negrito
start_bold = "\033[1m"
# Código ANSI para resetar a formatação
reset_bold = "\033[0m"

# listaemailsequipe = variaveis.emails


def process_exists(process_name):
    """
    :param process_name: nome do processo a ser verificado na lista de processos do Windows.
    :return: retorna a resposta se o programa está aberto (True) ou não (False).
    """
    import subprocess
    processes = \
        subprocess.Popen('tasklist', stdin=subprocess.PIPE, stderr=subprocess.PIPE,
                         stdout=subprocess.PIPE).communicate()[0]
    if process_name in str(processes):
        return True
    else:
        return False


def fecharprograma(nome):
    """

    :param nome: nome do executável a ser finalizado.
    """
    import os

    os.system("taskkill /im " + nome)


def caminhospadroes(caminho):
    """
    :param caminho: opção do caminho padrão que gostaria de retornar (em caso de dúvida ver lista abaixo).
    :return: o caminho segundo a opção dada como entrada.
    """
    import ctypes.wintypes
    # CSIDL	                        Decimal	Hex	    Shell	Description CSIDL_ADMINTOOLS	            48	    0x30	5.0	    The file system directory that is used to store administrative tools
    # for an individual user. CSIDL_ALTSTARTUP	            29	    0x1D	 	    The file system directory that corresponds to the user's nonlocalized Startup program group. CSIDL_APPDATA
    # 26	    0x1A	4.71	The file system directory that serves as a common repository for application-specific data. CSIDL_BITBUCKET	            10	    0x0A	 	    The virtual folder
    # containing the objects in the user's Recycle Bin. CSIDL_CDBURN_AREA	            59	    0x3B	6.0	    The file system directory acting as a staging area for files waiting to be written to
    # CD. CSIDL_COMMON_ADMINTOOLS	    47	    0x2F	5.0	    The file system directory containing administrative tools for all users of the computer. CSIDL_COMMON_ALTSTARTUP	    30
    # 0x1E	        NT-based only	The file system directory that corresponds to the nonlocalized Startup program group for all users. CSIDL_COMMON_APPDATA	        35	    0x23	5.0	    The
    # file system directory containing application data for all users. CSIDL_COMMON_DESKTOPDIRECTORY	25	    0x19	        NT-based only	The file system directory that contains files and
    # folders that appear on the desktop for all users. CSIDL_COMMON_DOCUMENTS	    46	    0x2E	 	    The file system directory that contains documents that are common to all users.
    # CSIDL_COMMON_FAVORITES	    31	    0x1F	        NT-based only	The file system directory that serves as a common repository for favorite items common to all users. CSIDL_COMMON_MUSIC
    # 53	    0x35	6.0	    The file system directory that serves as a repository for music files common to all users. CSIDL_COMMON_PICTURES	        54	    0x36	6.0	    The file system
    # directory that serves as a repository for image files common to all users. CSIDL_COMMON_PROGRAMS	        23	    0x17	        NT-based only	The file system directory that contains the
    # directories for the common program groups that appear on the Start menu for all users. CSIDL_COMMON_STARTMENU	    22	    0x16	        NT-based only	The file system directory that
    # contains the programs and folders that appear on the Start menu for all users. CSIDL_COMMON_STARTUP	        24	    0x18	        NT-based only	The file system directory that contains
    # the programs that appear in the Startup folder for all users. CSIDL_COMMON_TEMPLATES	    45	    0x2D	        NT-based only	The file system directory that contains the templates that are
    # available to all users. CSIDL_COMMON_VIDEO	        55	    0x37	6.0	    The file system directory that serves as a repository for video files common to all users.
    # CSIDL_COMPUTERSNEARME	        61	    0x3D	6.0	    The folder representing other machines in your workgroup. CSIDL_CONNECTIONS	            49	    0x31	6.0	    The virtual folder
    # representing Network Connections, containing network and dial-up connections. CSIDL_CONTROLS	            3	    0x03	 	    The virtual folder containing icons for the Control Panel
    # applications. CSIDL_COOKIES	                33	    0x21	 	    The file system directory that serves as a common repository for Internet cookies. CSIDL_DESKTOP	                0
    # 0x00	 	    The virtual folder representing the Windows desktop, the root of the shell namespace. CSIDL_DESKTOPDIRECTORY	    16	    0x10	 	    The file system directory used to
    # physically store file objects on the desktop. CSIDL_DRIVES	                17	    0x11	 	    The virtual folder representing My Computer, containing everything on the local computer:
    # storage devices, printers, and Control Panel. The folder may also contain mapped network drives. CSIDL_FAVORITES	            6	    0x06	 	    The file system directory that serves as a
    # common repository for the user's favorite items. CSIDL_FONTS	                20	    0x14	 	    A virtual folder containing fonts. CSIDL_HISTORY	                34	    0x22
    # The file system directory that serves as a common repository for Internet history items. CSIDL_INTERNET	            1	    0x01	 	    A viritual folder for Internet Explorer.
    # CSIDL_INTERNET_CACHE	        32	    0x20	4.72	The file system directory that serves as a common repository for temporary Internet files. CSIDL_LOCAL_APPDATA	        28	    0x1C
    # 5.0	    The file system directory that serves as a data repository for local (nonroaming) applications. CSIDL_MYDOCUMENTS	            5	    0x05	6.0	    The virtual folder
    # representing the My Documents desktop item. CSIDL_MYMUSIC	                13	    0x0D	6.0	    The file system directory that serves as a common repository for music files.
    # CSIDL_MYPICTURES	            39	    0x27	5.0	    The file system directory that serves as a common repository for image files. CSIDL_MYVIDEO	                14	    0x0E	6.0	    The
    # file system directory that serves as a common repository for video files. CSIDL_NETHOOD	                19	    0x13	 	    A file system directory containing the link objects that may
    # exist in the My Network Places virtual folder. CSIDL_NETWORK	                18	    0x12	 	    A virtual folder representing Network Neighborhood, the root of the network namespace
    # hierarchy. CSIDL_PERSONAL	            5	    0x05	 	    The file system directory used to physically store a user's common repository of documents. (From shell version 6.0 onwards,
    # CSIDL_PERSONAL is equivalent to CSIDL_MYDOCUMENTS, which is a virtual folder.) CSIDL_PHOTOALBUMS	            69	    0x45	Vista	The virtual folder used to store photo albums.
    # CSIDL_PLAYLISTS	            63	    0x3F	Vista	The virtual folder used to store play albums. CSIDL_PRINTERS	            4	    0x04	 	    The virtual folder containing
    # installed printers. CSIDL_PRINTHOOD	            27	    0x1B	 	    The file system directory that contains the link objects that can exist in the Printers virtual folder.
    # CSIDL_PROFILE	                40	    0x28	5.0	    The user's profile folder. CSIDL_PROGRAM_FILES	        38	    0x26	5.0	    The Program Files folder. CSIDL_PROGRAM_FILESX86
    # 42	    0x2A	5.0	    The Program Files folder for 32-bit programs on 64-bit systems. CSIDL_PROGRAM_FILES_COMMON	43	    0x2B	5.0	    A folder for components that are shared across
    # applications. CSIDL_PROGRAM_FILES_COMMONX86	44	    0x2C	5.0	    A folder for 32-bit components that are shared across applications on 64-bit systems. CSIDL_PROGRAMS	            2
    # 0x02	 	    The file system directory that contains the user's program groups (which are themselves file system directories). CSIDL_RECENT	                8	    0x08	 	    The file
    # system directory that contains shortcuts to the user's most recently used documents. CSIDL_RESOURCES	            56	    0x38	6.0	    The file system directory that contains resource data.
    # CSIDL_RESOURCES_LOCALIZED	    57	    0x39	6.0	    The file system directory that contains localized resource data. CSIDL_SAMPLE_MUSIC	        64	    0x40	Vista	The file system
    # directory that contains sample music. CSIDL_SAMPLE_PLAYLISTS	    65	    0x41	Vista	The file system directory that contains sample playlists. CSIDL_SAMPLE_PICTURES	        66
    # 0x42	Vista	The file system directory that contains sample pictures. CSIDL_SAMPLE_VIDEOS	        67	    0x43	Vista	The file system directory that contains sample videos.
    # CSIDL_SENDTO	                9	    0x09	 	    The file system directory that contains Send To menu items. CSIDL_STARTMENU	            11	    0x0B	 	    The file system directory
    # containing Start menu items. CSIDL_STARTUP	                7	    0x07	 	    The file system directory that corresponds to the user's Startup program group. CSIDL_SYSTEM
    # 37	    0x25	5.0	    The Windows System folder. CSIDL_SYSTEMX86	            41	    0x29	5.0	    The Windows 32-bit System folder on 64-bit systems. CSIDL_TEMPLATES	            21
    # 0x15	 	    The file system directory that serves as a common repository for document templates. CSIDL_WINDOWS	                36	    0x24	5.0	    The Windows directory or SYSROOT.

    csidl_personal = caminho  # Caminho padrão
    shgfp_type_current = 0  # Para não pegar a pasta padrão e sim a definida como documentos

    buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
    ctypes.windll.Shell32.SHGetFolderPathW(None, csidl_personal, None, shgfp_type_current, buf)

    return buf.value


def listaconexoes():
    """
    :return: vai ao arquivo xml padrão do SAP GUI para pegar todas as conexões instaladas no mesmo.
    """
    import xml.etree.ElementTree as ET

    servidores = []
    caminho = caminhospadroes(26) + '\\SAP\\Common\\SAPUILandscape.xml'
    if not os.path.isfile(caminho):
        caminho = 'C:\\ProgramData\\AutoPilotConfig\\SAP\\SAPUILandscape.xml'
        if not os.path.isfile(caminho):
            caminho = procurar_arquivo('SAPUILandscape.xml', 'C:\\')
    if os.path.isfile(caminho):
        mytree = ET.parse(caminho)
        myroot = mytree.getroot()
        for service in myroot.iter('Service'):
            servidores.append(service.get('name'))

    if len(servidores) > 0:
        return servidores
    else:
        return None


def criarinputbox(titulo, mensagem, substituircaracter='', valorinicial=''):
    """
    :param valorinicial: valor pré-preenchido na caixa de texto.
    :param titulo: cabeçalho da caixa de recebimento de dados do usuário (inputbox).
    :param mensagem: mensagem (normalmente descritiva ao 'input') para orientar o usuário.
    :param substituircaracter: caso seja um campo de senha informar o parâmetro para que a digitação não fique visível.
    :return: janela com as opções escolhidas.
    """
    import tkinter as tk
    from tkinter import simpledialog

    root = tk.Tk()
    root.attributes("-topmost", True)
    root.withdraw()

    # the input dialog
    user_inp = simpledialog.askstring(title=titulo, prompt=mensagem, initialvalue=valorinicial, show=substituircaracter)
    if user_inp is None:
        user_inp = 0

    return user_inp


def pesquisalista(lista, item):
    """
    :param lista: lista a ser realizado a busca.
    :param item: item a ser encontrado.
    :return: retorna o índice do item procurado ou -1 caso não ache.
    """
    try:
        return lista.index(item)

    except ValueError:
        return -1


def chunks(lista, n):
    """
    :param lista: lista a ser "quebrada".
    :param n: quantidade de itens por sublista
    """
    if len(lista) > 0:
        for i in range(0, len(lista), n):
            yield lista[i:i + n]


def retornarlista(dias=7, grupomail='', somentenaolidos=True):
    listadados = []
    textofiltro = ''

    # Conectar ao Outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # Selecionar a pasta de entrada (Inbox)
    if len(grupomail) == 0:
        inbox_folder = namespace.GetDefaultFolder(6)
    else:
        # Obtendo a pasta de entrada (Inbox) do grupo
        inbox_folder = outlook.GetDefaultFolder(6).Folders(grupomail)

    # Obter a coleção de e-mails
    emails = inbox_folder.Items
    # if dias > 0:
    #     data_limite = datetime.now() - timedelta(days=dias)
    #     textofiltro = "[ReceivedTime] >= '" + data_limite.strftime('%m/%d/%Y %H:%M %p') + "'"

    if somentenaolidos:
        if len(textofiltro) > 0:
            textofiltro = textofiltro + " AND [UnRead] = True"
        else:
            textofiltro = textofiltro + "[UnRead] = True"

    # Filtrar e-mails por data
    if len(textofiltro) > 0:
        emailsanalise = emails.Restrict(textofiltro)
    else:
        emailsanalise = emails

    emailsanalise.Sort("[ReceivedTime]", True)

    # last_email_id = None  # Armazenar o ID do último e-mail na cadeia

    for email in emailsanalise:

        # Exibir informações do e-mail
        if ('ZD12' in str(email.subject).upper() or 'ZD 12' in str(
                email.subject).upper()) and 'A ordem foi criada sob nº' not in email.body:
            anexovalido = False
            if email.Attachments.Count > 0:
                attachments = email.Attachments
                for attachment in attachments:
                    # Verifique se o anexo é um arquivo Excel (xlsx)
                    if attachment.FileName.endswith('.xlsx'):
                        # Crie um nome temporário para o arquivo
                        temp_file_path = os.path.join(os.getcwd(), attachment.FileName)
                        # Salve o anexo em um local temporário
                        attachment.SaveAsFile(temp_file_path)
                        # Leia o arquivo Excel em um DataFrame
                        excel_data = pd.read_excel(temp_file_path)
                        # Valida se a planilha em excel possui todos os campos
                        if valida_excel(excel_data, variaveis.camposdados):
                            # Objeto E-mail
                            excel_data['Email'] = email
                            # ID do E-mail
                            excel_data['ID Email'] = email.EntryID
                            # Adicione os dados ao DataFrame principal
                            listadados.extend(excel_data.to_dict(orient='records'))
                            # Se achou um anexo com dados válidos
                            anexovalido = True
                        # Remova o arquivo temporário
                        os.remove(temp_file_path)

            # Se não tiver um anexo válido ele vai buscar os dados no corpo do e-mail
            if not anexovalido:
                linha, itensinvalidos = extrair_campos_e_dados(email.Body)
                if email is not None:
                    linha['Email'] = email
                    linha['ID Email'] = email.EntryID
                    listadados.append(linha)
    if len(listadados) > 0:
        # Transforme a lista de dicionários em um DataFrame
        df = pd.DataFrame(listadados)
        # Verifique se cada valor na coluna está duplicado
        df = df.sort_values(by='ID Email')
        df['Duplicado'] = df.duplicated('ID Email', keep=False)
        listadados = df.to_dict(orient='records')

    return listadados


def extrair_campos_e_dados(email_body):
    padroes_regex = variaveis.regex

    # Encontrar a melhor parte correspondente ao padrão regex de cada linha para cada chave
    parte_regex_para_chave = encontrar_parte_regex(email_body, padroes_regex, variaveis.possiveis_separadores_descricao)

    # Retorna a lista de campos e os campos que não encontrou em duas listas separadas
    retorno, camposcomproblemas = parte_regex_para_chave

    de_para = variaveis.retornamailxtabela()

    novo_dicionario = {}
    # Criando um novo dicionário com as chaves atualizadas
    for chave, valor in retorno.items():
        if chave in de_para:
            novo_dicionario[de_para[chave]] = valor
        else:
            novo_dicionario[chave] = valor

    retorno = novo_dicionario

    # Transforma o dicionário em lista de dados simples
    # retorno = [retorno[item]['melhor_parte_regex'] for item in retorno]

    # Atualizando as chaves com base nos valores de 'melhor_parte_regex'
    retorno = {chave: valor['melhor_parte_regex'] for chave, valor in retorno.items()}

    return retorno, camposcomproblemas


def reply_or_forward_email(mail, reply_text, forward_to=None, textoaadicionar=None):
    signature = None
    if forward_to is not None:
        # If forward_to is provided, forward the email to the specified recipient.
        forward = mail.Forward()
        # Obtém a assinatura do Outlook
        # signature = pega_assinatura(0)
        if isinstance(forward_to, list):
            # If forward_to is a list of emails, join them with a semicolon as recipients.
            forward.To = ";".join(forward_to)
        else:
            # If forward_to is a single email, set it as the recipient.
            forward.To = forward_to

        if textoaadicionar is not None:
            forward.Body = "Erro: " + textoaadicionar + "\n\n" + forward.Body

        if signature is not None:
            forward.Body = forward.Body + "\n\n" + signature

        forward.Send()
        mail.UnRead = False
    else:
        # If forward_to is not provided, reply to the email with the given reply_text.
        reply = mail.ReplyAll()
        # signature = pega_assinatura(0)

        # reply.Body = reply_text
        if signature is not None:
            reply.Body = reply_text + "\n\n" + signature
        else:
            reply.Body = reply_text

        reply.Send()
        # Marcar o e-mail como lido
        mail.UnRead = False


def validadados(dicionario):
    retornodicionario = ''
    campoinvalido = False

    regex = variaveis.regex
    de_para = variaveis.retornamailxtabela()
    novo_retorno = {}
    for chave_original, chave_nova in de_para.items():
        if chave_original in regex:
            novo_retorno[chave_nova] = regex[chave_original]
        else:
            novo_retorno[chave_nova] = None

    regex = novo_retorno

    for campo, valor in dicionario.items():
        if campo in regex:
            if campo != 'Código do Cliente' or (campo == 'Código do Cliente' and len(valor.strip()) > 0):
                if not re.fullmatch(regex[campo], valor):
                    retornodicionario = retornodicionario + f"Campo: '{campo}' fora do padrão.\n"
                    campoinvalido = True

    if not campoinvalido:
        condicao, ordens = verificar_linhas("Lista Ordem Internas",
                                            dicionario['Centro de Lucro'],
                                            dicionario['Centro de Custo Responsável'],
                                            dicionario['Centro de Custo Solicitante'],
                                            dicionario['Área de Aplicação'],
                                            dicionario['Objetivo Setorial'],
                                            dicionario['Grupo de Ordem'],
                                            dicionario['Código do Cliente'])
        if condicao:
            campoinvalido = True
            retornodicionario = f'Combinação já existe, favor verificar! Ordens Disponíveis:{ordens}!'

    return campoinvalido, retornodicionario


def verificar_linhas(tabela, centro_lucro, centro_custos_responsavel,
                     centro_custos_solicitante, area_aplicacao, objetivo_setorial,
                     grupo_ordem, codcliente=''):
    # try:
    banco = bd.Banco()

    df = banco.consultar(f'SELECT * FROM [{tabela}]')  # Lê o arquivo Excel e cria um DataFrame

    if codcliente == '' or codcliente is None:
        # Verifica se a combinação de valores existe no DataFrame
        listaOI = df[(df['Centro de Lucro'] == centro_lucro) &
                     (df['Centro de Custo Responsável'] == centro_custos_responsavel) &
                     (df['Centro de Custo Solicitante'] == centro_custos_solicitante) &
                     (df['Área de Aplicação'] == area_aplicacao) &
                     (df['Objetivo Setorial'] == objetivo_setorial) &
                     (df['Grupo de Ordem'] == grupo_ordem)]
        combinacao_existe = listaOI.shape[0] > 0

    else:
        # Verifica se a combinação de valores existe no DataFrame
        listaOI = df[(df['Centro de Lucro'] == centro_lucro) &
                     (df['Centro de Custo Responsável'] == centro_custos_responsavel) &
                     (df['Centro de Custo Solicitante'] == centro_custos_solicitante) &
                     (df['Área de Aplicação'] == area_aplicacao) &
                     (df['Objetivo Setorial'] == objetivo_setorial) &
                     (df['Grupo de Ordem'] == grupo_ordem) &
                     (df['Código do Cliente'] == codcliente)]
        combinacao_existe = listaOI.shape[0] > 0

    return combinacao_existe, ','.join(listaOI['Ordem'].astype(str) + ', Status: ' + listaOI['Status'].astype(str))

    # except Exception as e:
    #     print(f"Ocorreu um erro ao ler o arquivo: {str(e.args[0])}")


def adicionardadosdf(arquivo, dfdadosnovos):
    # Carregar o DataFrame existente
    dfdadosnovos = dfdadosnovos[['Ordem',
                                 'Descrição da atividade',
                                 'Centro de Lucro',
                                 'Centro de Custos Responsável',
                                 'Centro de Custos Solicitante',
                                 'Área de Aplicação',
                                 'Objetivo Setorial',
                                 'Grupo de Ordem',
                                 'Código do Cliente',
                                 'Data Entrada',
                                 'Resposta']]

    df_existente = pd.read_excel(arquivo, sheet_name='Planilha1')

    # Anexar o novo DataFrame ao DataFrame existente
    df_final = pd.concat([df_existente, dfdadosnovos], ignore_index=True)

    # Escrever o DataFrame combinado no mesmo arquivo Excel
    with pd.ExcelWriter(arquivo, engine='openpyxl', mode='a') as writer:
        df_final.to_excel(writer, sheet_name='Planilha1', header=False, index=False, if_sheet_exists='replace')

    print('O novo DataFrame foi adicionado ao arquivo Excel existente.')


# Verifica se o dataframe tem a lista de campos dado como entrada
def valida_excel(df, campos_especificos):
    # Ajuste os nomes das colunas no DataFrame removendo espaços em branco extras
    df.columns = df.columns.str.strip()
    return all(campo.strip() in df.columns for campo in campos_especificos)


# Função para obter o conteúdo do arquivo de assinatura (boilerplate)
def get_boilerplate(file_path):
    try:
        # Tentar abrir o arquivo com a codificação UTF-8
        with open(file_path, 'r', encoding='utf-8') as file:
            return file.read()
    except UnicodeDecodeError:
        try:
            # Se a abertura com UTF-8 falhar, tentar abrir com a codificação Latin-1 (ISO-8859-1)
            with open(file_path, 'r', encoding='latin-1') as file:
                return file.read()
        except Exception as e:
            print("Erro ao obter o conteúdo do boilerplate:", e)
            return ""
    except Exception as e:
        print("Erro ao obter o conteúdo do boilerplate:", e)
        return ""


# Lista de arquivos .htm na pasta de assinaturas
def lista_arquivos_caminho(caminho, extensao):
    arquivos = []
    for arquivo in os.listdir(caminho):
        if arquivo.lower().endswith(extensao):
            arquivos.append(os.path.join(caminho, arquivo))
    return arquivos


def pega_assinatura(indice):
    sig_folder = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Signatures')
    # Obter a lista de arquivos .htm na pasta de assinaturas do Outlook
    arquivos_htm = lista_arquivos_caminho(sig_folder, ".htm")

    # Verificar se há pelo menos um arquivo .htm e obter o conteúdo do primeiro arquivo
    signature = ""
    if arquivos_htm:
        primeiro_arquivo_htm = arquivos_htm[indice]
        signature = get_boilerplate(primeiro_arquivo_htm)


def executar_bcp(tabela, arquivo_txt):
    nome_arquivo, extensao = os.path.splitext(arquivo_txt)
    bcp = caminhobcp()
    if verifica_caminho(arquivo_txt) == 'Nome do arquivo':
        diretorio_projeto = os.path.dirname(os.path.abspath(__file__))
        nome_arquivo = os.path.join(diretorio_projeto, nome_arquivo)
        arquivo_txt = os.path.join(diretorio_projeto, arquivo_txt)

    comando = f'"{bcp}" {pwd.schema}."[{tabela}]" IN "{arquivo_txt}" -t "|" -C SQL_Latin1_General_CP1_CI_AS -c -S {pwd.endbanco} -U {pwd.usrbanco} -P {pwd.pwdbanco} -d {pwd.nomebanco} -e "{nome_arquivo}.err" -F 2 > "{nome_arquivo}.log"'
    proc = subprocess.Popen(comando, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    stdout, stderr = proc.communicate()

    stdout_text = decode_byte_to_text(stdout, 'ANSI')
    stderr_text = decode_byte_to_text(stderr, 'ANSI')

    if stdout is not None and stderr is not None:
        if 'Error' in stdout_text:
            msg.msgbox(
                f'Falha na carga do arquivo:\n {arquivo_txt} \nDescrição Erro:\n {decode_byte_to_text(stdout, "ANSI")}',
                msg.MB_OK, 'Erro BCP!')
            os.system('notepad ' + arquivo_txt)
            return None

        if stderr_text:
            print(decode_byte_to_text(stderr, "ANSI"))
            msg.msgbox(
                f'Falha na carga do arquivo:\n {arquivo_txt} \nDescrição Erro:\n {decode_byte_to_text(stderr, "ANSI")}',
                msg.MB_OK, 'Erro BCP!')
            os.system('notepad ' + arquivo_txt)
            return None

    num_rows_inserted = 0
    arquivolog = f"{nome_arquivo}.log"
    if os.path.exists(arquivolog):
        with open(arquivolog, 'r') as log:
            for line in log:
                if 'linhas copiadas' in line:
                    num_rows_inserted = int(line.split()[0])
                    break

    return num_rows_inserted


def caminhobcp():
    import shutil

    # Tenta encontrar o caminho do executável BCP no PATH
    bcp_path = shutil.which("bcp")
    return bcp_path


def decode_byte_to_text(entrada, decoder=''):
    import chardet

    if decoder == '':
        detected_encoding = chardet.detect(entrada)['encoding']
        if detected_encoding is not None:
            decoded_output = entrada.decode(detected_encoding)
        else:
            decoded_output = ''
    else:
        decoded_output = entrada.decode(decoder)
    return decoded_output


def verifica_caminho(string):
    nome_arquivo = os.path.basename(string)

    # Se a string for igual ao nome do arquivo, então é apenas o nome do arquivo
    if string == nome_arquivo:
        return "Nome do arquivo"
    else:
        return "Caminho completo"


def arquivotodf(arquivo):
    # Ler o arquivo de texto
    with open(arquivo, 'r') as file:
        lines = file.readlines()

    # Encontrar a linha do cabeçalho
    header_line = next((i for i, line in enumerate(lines) if '|' in line), None)

    # Se não houver linha com '|', sair
    if header_line is None:
        print("Não foi encontrada uma linha com cabeçalhos válidos.")
        exit()

    # Extrair os cabeçalhos da tabela
    header = [item.strip() for item in lines[header_line].split('|') if item.strip()]

    # Extrair os dados da tabela
    data = []
    for line in lines[header_line + 1:]:
        items = line.split('|')
        # Verificar se o número de itens é igual ao número de cabeçalhos
        print(len(items), len(header))
        if len(items) - 2 == len(header):
            items = line.split('|')[1:-1]
            data_row = []
            for item in items:
                # Se o campo for vazio, adicionar um espaço em branco, senão adicionar o valor
                data_row.append(item.strip() if item.strip() else None)
            data.append(data_row)

    # Criar DataFrame
    df = pd.DataFrame(data, columns=header)

    # Exibir DataFrame
    return df


def tratararquivo(nomearquivo, num_colunas=0, retornarcabecalho=False, retornadf=False):
    linhascompletas = []
    achoucabecalho = False
    linhacompleta = ""
    cabecalho = ""
    column_names = []

    with open(nomearquivo, 'r') as file:
        for line in file:
            line = line.strip()
            if line.startswith("|") or line.startswith("|"):
                if line.startswith("|") and line.startswith("|"):
                    line = line[1:-1].strip()
                    provisorio = line.split("|")
                    if line != len(line) * '-' and len(line) > 0 and provisorio[0].strip() != '*':
                        if num_colunas == 0:
                            num_colunas = len(provisorio)
                        if len(provisorio) == num_colunas:
                            if not achoucabecalho:
                                cabecalho = line.strip()
                                linhacompleta = line.strip()
                                column_names = [field.strip() for field in provisorio]
                                achoucabecalho = True
                            else:
                                if len(cabecalho) > 0:
                                    if line != cabecalho:
                                        linhacompleta = line
                else:
                    if num_colunas > 0:
                        if linhacompleta != "" and line.endswith("|") and len(
                                str(linhacompleta + line)[1:-1].split("|")) <= num_colunas:
                            linhacompleta += line[1:-1]
                        else:
                            if not (len(str(linhacompleta + line)[1:-1].split("|")) <= num_colunas):
                                linhacompleta = ""
            else:
                if len(str(linhacompleta + line)[1:-1].split("|")) <= num_colunas and line != len(line) * '-':
                    linhacompleta += line[1:-1]
                else:
                    linhacompleta = ""

            provisorio = linhacompleta.split("|")
            if len(provisorio) == num_colunas and linhacompleta != cabecalho and line != len(line) * '-':
                linhascompletas.append([field.strip() for field in provisorio])
                linhacompleta = ''

            provisorio = []
    if not retornadf:
        if not retornarcabecalho:
            return linhascompletas

        else:
            return linhascompletas, column_names
    else:
        if linhascompletas:
            # Concatenar as listas internas em uma única lista
            # lista_total = [item for sublist in linhascompletas for item in sublist]
            # Concatenar as listas internas em uma única lista, excluindo o cabeçalho da segunda lista
            # lista_total = [item for i, sublist in enumerate(listadedadosdoarquivo) for item in sublist if i == 0 or i > 0 and sublist.index(item) > 0]
            dataframe = pd.DataFrame(linhascompletas, columns=column_names)
            return dataframe


def procurar_arquivo(nome_arquivo, diretorio):
    # Percorre recursivamente os diretórios

    # Registrar o tempo de início
    inicio = time.time()

    for root, dirs, files in os.walk(diretorio):
        if nome_arquivo in files:
            return os.path.join(root, nome_arquivo)

    # Registrar o tempo de fim
    fim = time.time()

    # Calcular o tempo decorrido em segundos
    tempo_decorrido = fim - inicio
    print("Tempo decorrido:", tempo_decorrido, "segundos")

    # Se o arquivo não for encontrado
    return None


# Função para encontrar a melhor parte correspondente ao padrão regex de cada linha para cada chave
def encontrar_parte_regex(texto, padroes_regex, possiveis_separadores, limiteminimo=50):
    obrigatoriosseparador = ['Descrição da atividade (Ordem Interna)',
                             'Área de Aplicação (programa de investimentos)']
    linhaadesconsiderar = []
    parte_regex_para_chave = {}
    for chave, padrao_regex in padroes_regex.items():
        if 'Área de Aplicação' in chave or 'Descrição da atividade' in chave:
            w = 1
        linha_regex = ''
        melhor_parte_regex = ''
        melhor_score = 0
        for linha in texto.split('\n'):
            linha = linha.strip()
            if linha not in linhaadesconsiderar:
                match = re.search(padrao_regex, linha)
                if match:
                    score = fuzz.partial_ratio(chave, linha)
                    if score > melhor_score:
                        melhor_score = score
                        if score > limiteminimo:
                            if chave != 'Descrição da atividade (Ordem Interna)':
                                melhor_parte_regex = match.group().strip()
                                linha_regex = linha
                            else:
                                # Encontrar a parte após o primeiro separador encontrado
                                for separador in possiveis_separadores:
                                    indice_separador = linha.find(separador)
                                    if indice_separador != -1:
                                        melhor_parte_regex = linha[indice_separador + 1:].strip()
                                        linha_regex = linha
                                        break
                                    else:
                                        melhor_parte_regex = ''  # Se não houver separador, retorna nada

                        else:
                            melhor_parte_regex = ''
                            linha_regex = linha

        if melhor_score > limiteminimo and linha not in linhaadesconsiderar:
            linhaadesconsiderar.append(linha_regex)
        parte_regex_para_chave[chave] = {'melhor_parte_regex': melhor_parte_regex,
                                         'melhor_score': melhor_score}  # 'linha_score': linha_regex,

    # Verificar chaves sem item
    chaves_sem_item = [chave for chave, valor in parte_regex_para_chave.items() if not valor['melhor_parte_regex']]

    return parte_regex_para_chave, chaves_sem_item


def remover_caracteres_ilegais(texto):
    caracteres_permitidos = (
            ''.join(chr(i) for i in range(32, 126)) +
            ''.join(chr(i) for i in range(160, 255))
    )
    return ''.join(c for c in texto if c in caracteres_permitidos)


def formatar_data(data):
    if data.hour == 0 and data.minute == 0 and data.second == 0:
        return data.strftime('%Y/%m/%d')
    else:
        return data.strftime('%Y/%m/%d %H:%M:%S')
