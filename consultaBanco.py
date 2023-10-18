import pyodbc
from sqlalchemy import create_engine

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
