# Separadores possíveis
possiveis_separadores_descricao = [':', '-', '/', '|', '\t']
separadores = ''.join(possiveis_separadores_descricao)
# Dicionário com os padrões regex
regex = {
    'Descrição da atividade (Ordem Interna)': f'({separadores}|\\s+|^).*?(?=[ \r\n]|$)',
    'Centro de Lucro': '[A-Za-z]{3}[A-Za-z0-9]{4}[0-9]{3}(?=[ \r\n]|$)',
    'Centro de Custo Responsável': '[A-Za-z]{3}[A-Za-z0-9]{4}[0-9]{3}(?=[ \r\n]|$)',
    'Centro de Custo Solicitante': '[A-Za-z]{3}[A-Za-z0-9]{4}[0-9]{3}(?=[ \r\n]|$)',
    'Área de Aplicação (programa de investimentos)': f'({separadores}|\\s+|^)[A-Za-z]{{4}}(?=[ \r\n]|$)',
    'Objetivo Setorial (PRJ)': 'PRJ[0-9]{8}(?=[ \r\n]|$)',
    'Grupo de Ordem (PI)': 'PI[0-9]{3}(?=[ \r\n]|$)',
    'Código do Cliente, se aplicável': '[0-9]+(?=[ \r\n]|$)'
}

camposdados = ['Descrição da atividade',
               'Centro de Lucro',
               'Centro de Custo Responsável',
               'Centro de Custo Solicitante',
               'Área de Aplicação',
               'Objetivo Setorial',
               'Grupo de Ordem',
               'Código do Cliente']


emails = ['jessicaabreu@vibraenergia.com.br',
          'rodricastro@vibraenergia.com.br',
          'iriscoelho@vibraenergia.com.br',
          'ericksperle@vibraenergia.com.br']

def retornamailxtabela():
    listamailxtabela = {}
    for para, regex_para in regex.items():
        for de in camposdados:
            if de.lower() in para.lower():
                listamailxtabela[para] = de
                break  # Para evitar iterações desnecessárias
    return listamailxtabela

