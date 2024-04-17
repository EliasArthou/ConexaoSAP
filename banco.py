import os
import pyodbc
import sensiveis as senha
import pandas as pd
from sqlalchemy import create_engine
from urllib.parse import quote_plus
import auxiliares as aux
import sys

codentrada = 'ANSI'


class Banco:
    """
    Criado para se conectar e realizar operações no banco de dados
    """

    def __init__(self):
        self.conxn = None
        self.cursor = None
        self.erro = ''
        self.engine = None
        self.constr = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + senha.endbanco + ';DATABASE=' + senha.nomebanco + ';UID=' + senha.usrbanco + ';PWD=' + senha.pwdbanco
        # Crie a string de conexão
        self.params = quote_plus(self.constr)

        self.abrirconexao()

    def abrirconexao(self):
        try:
            if len(self.constr) > 0:
                self.engine = create_engine("mssql+pyodbc:///?odbc_connect=%s" % self.params)
                self.conxn = pyodbc.connect(self.constr)
                self.cursor = self.conxn.cursor()
                self.cursor.execute("SELECT 1")
                self.cursor.close()
                self.cursor = self.conxn.cursor()

        except pyodbc.Error as e:
            self.erro = str(e)

    def consultar(self, sql):
        """
        :param sql: Código sql a ser executado (uma consulta SQL).
        :return: O resultado da consulta em uma lista.
        """
        if not (self.conxn is not None and not self.conxn.closed):
            self.abrirconexao()

        resultado = pd.read_sql(sql, self.engine)
        return resultado

    def adicionardfemmassa(self, tabela, df, anomes, campos_where=[], indices_filtro=[], viabcp=True):
        try:
            linhascarregadas = 0
            tabelaexiste = self.verificartabela(tabela)

            if tabelaexiste and len(campos_where) == len(indices_filtro):
                if len(campos_where) != 0:
                    if campos_where and indices_filtro:
                        where_conditions = " AND ".join([f"{coluna} = '{df.iloc[:, indice].unique()[0]}'" for coluna, indice in zip(campos_where, indices_filtro)])
                        strSQL = f"DELETE FROM {tabela} WHERE {where_conditions}"
                        if self.conxn is None or self.conxn.closed:
                            self.abrirconexao()
                        self.cursor.execute(strSQL)
                        self.conxn.commit()
                if viabcp:
                    # Extensões do arquivo de entrada que deve ser excluído (a entrada é o TXT)
                    extensoes = ['err', 'log', 'txt']
                    # "Caminho" do arquivo que ficará a saída
                    arquivosaida = anomes.replace('/', '_') + '.txt'
                    # Caminho do Arquivo
                    caminho_pasta = os.path.dirname(arquivosaida)
                    # Separando o arquivo da extensão
                    nome_arquivo, extensao_txt = os.path.splitext(os.path.basename(arquivosaida))

                    # Looping paa excluir os arquivos das extensões informadas
                    for extensao in extensoes:
                        extensao_lower = extensao.lower()  # Converte para minúsculas
                        caminho_arquivo_atual = os.path.join(caminho_pasta, f"{nome_arquivo}.{extensao_lower}")
                        if os.path.exists(caminho_arquivo_atual):
                            os.remove(caminho_arquivo_atual)
                    if not os.path.exists(arquivosaida):
                        df.to_csv(arquivosaida, index=None, sep='|', mode='a', encoding=codentrada)
                        if os.path.isfile(arquivosaida):
                            linhascarregadas = aux.executar_bcp(tabela, arquivosaida)
                else:
                    df.columns = self.rename_duplicate_columns(df)

                    self.engine = create_engine("mssql+pyodbc:///?odbc_connect=%s" % self.params)
                    if not df.empty:
                        df.to_sql(tabela, con=self.engine, index=False, if_exists='append', method='multi', chunksize=90)

                    # Faça o commit após todas as operações

            return linhascarregadas
        except pyodbc.Error as e:
            # Em caso de erro, faça rollback
            self.conxn.rollback()

            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = exc_tb.tb_frame.f_code.co_filename
            line_number = exc_tb.tb_lineno
            error_message = f"Erro na linha {line_number} do arquivo {fname}: {e}"
            print("Erro ao realizar operações:", error_message)

        finally:
            # Certifique-se de fechar a conexão ao final
            if self.conxn is not None:
                self.conxn.close()

    def adicionardf(self, tabela, df, lista=[]):
        # try:
        tabelaexiste = self.verificartabela(tabela)
        if tabelaexiste:
            if self.conxn is not None and not self.conxn.closed:
                self.conxn.close()

            self.abrirconexao()
            # if len(lista) > 0:
            #     df = df.drop(columns=lista)
            listacampos = ['Ordem',
                           'Descrição da atividade',
                           'Centro de Lucro',
                           'Centro de Custo Responsável',
                           'Centro de Custo Solicitante',
                           'Área de Aplicação',
                           'Objetivo Setorial',
                           'Grupo de Ordem',
                           'Código do Cliente',
                           'Data Entrada',
                           'Status']
            df = df[listacampos]
            self.engine = create_engine("mssql+pyodbc:///?odbc_connect=%s" % self.params)
            df.to_sql(tabela, self.engine, index=False, if_exists='append')

        # except pyodbc.Error as e:
        #     # Em caso de erro, faça rollback
        #     self.conxn.rollback()
        #
        #     exc_type, exc_obj, exc_tb = sys.exc_info()
        #     fname = exc_tb.tb_frame.f_code.co_filename
        #     line_number = exc_tb.tb_lineno
        #     error_message = f"Erro na linha {line_number} do arquivo {fname}: {e}"
        #     print("Erro ao realizar operações:", error_message)
        #
        # finally:
        #     # Certifique-se de fechar a conexão ao final
        #     self.conxn.close()

    def verificartabela(self, tabela):
        try:
            if not (self.conxn is not None and not self.conxn.closed):
                self.abrirconexao()

            # Verifica se a tabela existe
            check_table_query = f"SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{tabela}'"
            cursorlocal = self.conxn.cursor()
            cursorlocal.execute(check_table_query)
            table_exists = cursorlocal.fetchone()[0] > 0
            cursorlocal.close()

            return table_exists

        finally:
            self.conxn.close()
            self.conxn = None

    # Função para renomear colunas duplicadas
    def rename_duplicate_columns(self, df):
        seen_cols = {}
        new_columns = []

        for col in df.columns:
            if col in seen_cols:
                seen_cols[col] += 1
                new_columns.append(f"{col}{seen_cols[col]}")
            else:
                seen_cols[col] = 0
                new_columns.append(col)

        return new_columns

    def executarSQL(self, sql):
        try:
            if self.conxn is None or self.conxn.closed:
                self.abrirconexao()
            self.cursor.execute(sql)
            rows_affected = self.cursor.rowcount
            self.conxn.commit()
            return rows_affected
        finally:
            # Certifique-se de fechar a conexão ao final
            self.conxn.close()
