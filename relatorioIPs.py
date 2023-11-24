import pandas as pd
from sqlalchemy import create_engine
import pyodbc
import pygetwindow as gw
from pygetwindow import getWindowsWithTitle
from ping3 import ping

# Função para estabelecer a conexão com o banco de dados
def conectar_banco():
	# Configurações de conexão com o banco de dados
    server = 'CORSVMBDSPPDV01'
    database = 'VM_DataBSP'
    username = 'UserMDC'
    password = '@mdc2015@'

    connection_string = f'DRIVER=SQL Server;SERVER={server};DATABASE={database};UID={username};PWD={password}'

    try:
        connection = pyodbc.connect(connection_string)
        return connection, connection_string  # Retorna a conexão e a string de conexão
    except Exception as e:
        print(f"Erro na conexão com o banco de dados: {str(e)}")
        return None, None


# Função para realizar a consulta SQL e salvar em um arquivo Excel usando SQLAlchemy
def consultar_e_salvar_em_excel(connection, connection_string):
    if connection is None:
        return

    try:
        # Consulta SQL usando SQLAlchemy
        query = """
        SELECT L.CODIGO, L.DESCRICAO, C.LOCALIZACAO, C.IP
        FROM lojas L
        JOIN [VM_DataBSP].[dbo].[COMPONENTES] C ON L.CODIGO = C.LOJA
        WHERE L.CODIGO = 5137
        AND C.IP <> ''
        AND C.TIPO <> 'D'
        """

        # Criar uma conexão compatível com o pandas usando SQLAlchemy
        engine = create_engine("mssql+pyodbc:///?odbc_connect=" + connection_string)

        # Ler os resultados da consulta SQL em um DataFrame do pandas
        df = pd.read_sql_query(query, engine)

        # Caminho do arquivo Excel
        excel_file_path = r'C:\24 - RPA\Atualizacao_PDV_Py\Banco\IPs_PDVs.xlsx'

        # Salvar o DataFrame como um arquivo Excel
        df.to_excel(excel_file_path, index=False)

        print(f'Dados salvos em {excel_file_path}')

    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")

    finally:
        # Fechar a conexão com o banco de dados usando pyodbc
        connection.close()
    
def testarPing(excel_file_path):
    try:
        # Ler o arquivo Excel
        df = pd.read_excel(excel_file_path)

        # Realizar o teste de ping e adicionar os resultados ao DataFrame
        def ping_ip(ip):
            result = ping(ip)
            if result is not None:
                return "Ping OK"
            return "Falha no Ping"
        
        df['Ping Result'] = df['IP'].apply(ping_ip)

        # Salvar o DataFrame atualizado no mesmo arquivo Excel
        df.to_excel(excel_file_path, index=False, engine='openpyxl')

        print(f'Teste de ping concluído e dados atualizados em {excel_file_path}')

    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")

# Caminho do arquivo Excel
excel_file_path = r'C:\24 - RPA\Atualizacao_PDV_Py\Banco\IPs_PDVs.xlsx'

# Caminho da imagem do botão de conexão
imagem_path = r'C:\24 - RPA\Atualizacao_PDV_Py\Banco\botao_conexao.png'

# Chame a função conectar_banco para obter a conexão e a string de conexão
conexao, connection_string = conectar_banco()
# Verifique se a conexão foi estabelecida com sucesso antes de continuar
if conexao is not None:
    # Chame a função consultar_e_salvar_em_excel passando a conexão e a string de conexão
    consultar_e_salvar_em_excel(conexao, connection_string)

    # Chame a função testarPing passando o caminho do arquivo Excel

testarPing(excel_file_path)
