"""
Funções complementares de ações gerais para auxiliar o funcionamento do processo.
"""
import pypyodbc
import sensiveis as pwd


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
    # CSIDL	                        Decimal	Hex	    Shell	Description
    # CSIDL_ADMINTOOLS	            48	    0x30	5.0	    The file system directory that is used to store administrative tools for an individual user.
    # CSIDL_ALTSTARTUP	            29	    0x1D	 	    The file system directory that corresponds to the user's nonlocalized Startup program group.
    # CSIDL_APPDATA	                26	    0x1A	4.71	The file system directory that serves as a common repository for application-specific data.
    # CSIDL_BITBUCKET	            10	    0x0A	 	    The virtual folder containing the objects in the user's Recycle Bin.
    # CSIDL_CDBURN_AREA	            59	    0x3B	6.0	    The file system directory acting as a staging area for files waiting to be written to CD.
    # CSIDL_COMMON_ADMINTOOLS	    47	    0x2F	5.0	    The file system directory containing administrative tools for all users of the computer.
    # CSIDL_COMMON_ALTSTARTUP	    30	    0x1E	        NT-based only	The file system directory that corresponds to the nonlocalized Startup program group for all users.
    # CSIDL_COMMON_APPDATA	        35	    0x23	5.0	    The file system directory containing application data for all users.
    # CSIDL_COMMON_DESKTOPDIRECTORY	25	    0x19	        NT-based only	The file system directory that contains files and folders that appear on the desktop for all users.
    # CSIDL_COMMON_DOCUMENTS	    46	    0x2E	 	    The file system directory that contains documents that are common to all users.
    # CSIDL_COMMON_FAVORITES	    31	    0x1F	        NT-based only	The file system directory that serves as a common repository for favorite items common to all users.
    # CSIDL_COMMON_MUSIC	        53	    0x35	6.0	    The file system directory that serves as a repository for music files common to all users.
    # CSIDL_COMMON_PICTURES	        54	    0x36	6.0	    The file system directory that serves as a repository for image files common to all users.
    # CSIDL_COMMON_PROGRAMS	        23	    0x17	        NT-based only	The file system directory that contains the directories for the common program groups that appear on the Start menu for all users.
    # CSIDL_COMMON_STARTMENU	    22	    0x16	        NT-based only	The file system directory that contains the programs and folders that appear on the Start menu for all users.
    # CSIDL_COMMON_STARTUP	        24	    0x18	        NT-based only	The file system directory that contains the programs that appear in the Startup folder for all users.
    # CSIDL_COMMON_TEMPLATES	    45	    0x2D	        NT-based only	The file system directory that contains the templates that are available to all users.
    # CSIDL_COMMON_VIDEO	        55	    0x37	6.0	    The file system directory that serves as a repository for video files common to all users.
    # CSIDL_COMPUTERSNEARME	        61	    0x3D	6.0	    The folder representing other machines in your workgroup.
    # CSIDL_CONNECTIONS	            49	    0x31	6.0	    The virtual folder representing Network Connections, containing network and dial-up connections.
    # CSIDL_CONTROLS	            3	    0x03	 	    The virtual folder containing icons for the Control Panel applications.
    # CSIDL_COOKIES	                33	    0x21	 	    The file system directory that serves as a common repository for Internet cookies.
    # CSIDL_DESKTOP	                0	    0x00	 	    The virtual folder representing the Windows desktop, the root of the shell namespace.
    # CSIDL_DESKTOPDIRECTORY	    16	    0x10	 	    The file system directory used to physically store file objects on the desktop.
    # CSIDL_DRIVES	                17	    0x11	 	    The virtual folder representing My Computer, containing everything on the local computer: storage devices, printers, and Control Panel. The folder may also contain mapped network drives.
    # CSIDL_FAVORITES	            6	    0x06	 	    The file system directory that serves as a common repository for the user's favorite items.
    # CSIDL_FONTS	                20	    0x14	 	    A virtual folder containing fonts.
    # CSIDL_HISTORY	                34	    0x22	 	    The file system directory that serves as a common repository for Internet history items.
    # CSIDL_INTERNET	            1	    0x01	 	    A viritual folder for Internet Explorer.
    # CSIDL_INTERNET_CACHE	        32	    0x20	4.72	The file system directory that serves as a common repository for temporary Internet files.
    # CSIDL_LOCAL_APPDATA	        28	    0x1C	5.0	    The file system directory that serves as a data repository for local (nonroaming) applications.
    # CSIDL_MYDOCUMENTS	            5	    0x05	6.0	    The virtual folder representing the My Documents desktop item.
    # CSIDL_MYMUSIC	                13	    0x0D	6.0	    The file system directory that serves as a common repository for music files.
    # CSIDL_MYPICTURES	            39	    0x27	5.0	    The file system directory that serves as a common repository for image files.
    # CSIDL_MYVIDEO	                14	    0x0E	6.0	    The file system directory that serves as a common repository for video files.
    # CSIDL_NETHOOD	                19	    0x13	 	    A file system directory containing the link objects that may exist in the My Network Places virtual folder.
    # CSIDL_NETWORK	                18	    0x12	 	    A virtual folder representing Network Neighborhood, the root of the network namespace hierarchy.
    # CSIDL_PERSONAL	            5	    0x05	 	    The file system directory used to physically store a user's common repository of documents. (From shell version 6.0 onwards, CSIDL_PERSONAL is equivalent to CSIDL_MYDOCUMENTS, which is a virtual folder.)
    # CSIDL_PHOTOALBUMS	            69	    0x45	Vista	The virtual folder used to store photo albums.
    # CSIDL_PLAYLISTS	            63	    0x3F	Vista	The virtual folder used to store play albums.
    # CSIDL_PRINTERS	            4	    0x04	 	    The virtual folder containing installed printers.
    # CSIDL_PRINTHOOD	            27	    0x1B	 	    The file system directory that contains the link objects that can exist in the Printers virtual folder.
    # CSIDL_PROFILE	                40	    0x28	5.0	    The user's profile folder.
    # CSIDL_PROGRAM_FILES	        38	    0x26	5.0	    The Program Files folder.
    # CSIDL_PROGRAM_FILESX86	    42	    0x2A	5.0	    The Program Files folder for 32-bit programs on 64-bit systems.
    # CSIDL_PROGRAM_FILES_COMMON	43	    0x2B	5.0	    A folder for components that are shared across applications.
    # CSIDL_PROGRAM_FILES_COMMONX86	44	    0x2C	5.0	    A folder for 32-bit components that are shared across applications on 64-bit systems.
    # CSIDL_PROGRAMS	            2	    0x02	 	    The file system directory that contains the user's program groups (which are themselves file system directories).
    # CSIDL_RECENT	                8	    0x08	 	    The file system directory that contains shortcuts to the user's most recently used documents.
    # CSIDL_RESOURCES	            56	    0x38	6.0	    The file system directory that contains resource data.
    # CSIDL_RESOURCES_LOCALIZED	    57	    0x39	6.0	    The file system directory that contains localized resource data.
    # CSIDL_SAMPLE_MUSIC	        64	    0x40	Vista	The file system directory that contains sample music.
    # CSIDL_SAMPLE_PLAYLISTS	    65	    0x41	Vista	The file system directory that contains sample playlists.
    # CSIDL_SAMPLE_PICTURES	        66	    0x42	Vista	The file system directory that contains sample pictures.
    # CSIDL_SAMPLE_VIDEOS	        67	    0x43	Vista	The file system directory that contains sample videos.
    # CSIDL_SENDTO	                9	    0x09	 	    The file system directory that contains Send To menu items.
    # CSIDL_STARTMENU	            11	    0x0B	 	    The file system directory containing Start menu items.
    # CSIDL_STARTUP	                7	    0x07	 	    The file system directory that corresponds to the user's Startup program group.
    # CSIDL_SYSTEM	                37	    0x25	5.0	    The Windows System folder.
    # CSIDL_SYSTEMX86	            41	    0x29	5.0	    The Windows 32-bit System folder on 64-bit systems.
    # CSIDL_TEMPLATES	            21	    0x15	 	    The file system directory that serves as a common repository for document templates.
    # CSIDL_WINDOWS	                36	    0x24	5.0	    The Windows directory or SYSROOT.

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


class Conec:
    """
    Objeto de conexão com o banco de dados.
    """
    string: ''

    def __init__(self):
        self.string = "DRIVER={SQL Server};SERVER=" + pwd.endbanco + ";UID=" + pwd.usrbanco + ";PWD=" \
                      + pwd.pwdbanco + ";DATABASE=" + pwd.nomebanco

    def consulta(self, query, dictionary=False):
        """
        :param query: consulta a ser executada no banco (SELECT).
        :param dictionary: se vai retornar em forma de dicionário ou lista. O padrão é lista.
        :return: a lista ou dicionario com o resultado da consulta.
        """

        connection = pypyodbc.connect(self.string)
        cursor = connection.cursor()
        cursor.execute(query)
        if not dictionary:
            teste = cursor.fetchall()
            if len(teste[0]) != 1:
                results = teste.encode('ascii', 'ignore')
            else:
                results = [item[0] for item in teste]
        else:
            columns = [column[0] for column in cursor.description]
            results = []
            for row in cursor.fetchall():
                results.append(dict(zip(columns, row)))

        return results


def chunks(lista, n):
    """
    :param lista: lista a ser "quebrada".
    :param n: quantidade de itens por sublista
    """
    if len(lista) > 0:
        for i in range(0, len(lista), n):
            yield lista[i:i + n]


def list_to_clipboard(output_list, item=0):
    """
    Check if len(list) > 0, then copy to clipboard in dataframe format (text has problem to paste in SAP)

    :param output_list: lista que vai pra memória
    :param item: sublista copiada (controle meu)
    """

    import pandas as pd
    from ctypes import windll

    if item == 0:
        itemcopiado = ''
    else:
        itemcopiado = str(item)

    # Limpar o clipboard
    if windll.user32.OpenClipboard(None):
        windll.user32.EmptyClipboard()
        windll.user32.CloseClipboard()

    # Pegar o pedaço recebido da lista, transformar em dataframe e "jogar" para o ‘clipboard’ do Windows
    if len(output_list) > 0:
        df = pd.DataFrame(output_list, columns=[''])

        df.to_clipboard(True, '\r', index=False)

        print('Lista Copiada ' + str(itemcopiado))
    else:
        print('Sem itens na lista para copiar ' + str(itemcopiado))
