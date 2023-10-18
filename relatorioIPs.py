import pandas as pd
from sqlalchemy import create_engine

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
        WHERE L.CODIGO BETWEEN 5001 AND 5300
        AND C.IP <> ''
        AND C.TIPO <> 'D'
        AND C.FUNCAO NOT IN (1)
        ORDER BY CODIGO, LOCALIZACAO ASC
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