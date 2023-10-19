import pandas as pd
from sqlalchemy import create_engine
import pyodbc
import time
import pyautogui
import pygetwindow as gw
import subprocess
from pygetwindow import getWindowsWithTitle

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

def abrirTrauma():
    # Especifique o caminho completo para o executável do programa
    program_path = r'C:\Program Files (x86)\Trauma Zer0\Tz0\ModulesApp\Tz0Console.exe'

    # Iniciar o programa
    subprocess.Popen(program_path)

    # Aguarde o programa carregar (ajuste o tempo conforme necessário)
    time.sleep(20)

    # Verifique se a janela de autenticação está presente
    authentication_window = None
    for window in gw.getAllTitles():
        if "42725994888-OTISGK" in window:
            authentication_window = gw.getWindowsWithTitle(window)[0]
            break

    # Se a janela de autenticação for encontrada
    if authentication_window:
        # Clique no campo de senha
        x, y = 728, 289  # Substitua pelas coordenadas reais
        pyautogui.click(x, y)

        # Digite a senha
        pyautogui.typewrite("Mudar@2023")

        # Pressione Enter (ou a tecla de envio apropriada)
        pyautogui.press('enter')
    else:
        print("OK")

def find_or_open_tz0_console(program_path):
    # Verifique se a janela do Tz0Console já está aberta
    tz0_console_window = getWindowsWithTitle('Tz0 Console')
    if not tz0_console_window:
        # Se não estiver aberta, inicie o programa
        subprocess.Popen(program_path)
        time.sleep(20)

    return getWindowsWithTitle('Tz0 Console')[0]

def conectaIP(excel_file_path):
    # Ler o arquivo Excel
    df = pd.read_excel(excel_file_path)

    # Filtrar os IPs com "Ping OK"
    df_ping_ok = df[df['Ping Result'] == 'Ping OK']

    # Obter a lista de IPs filtrados
    ips_ping_ok = df_ping_ok['IP'].tolist()

    # Escolha o IP a ser conectado (por exemplo, o primeiro IP da lista)
    if ips_ping_ok:
        ip_to_connect = ips_ping_ok[0]

        # Encontre o código e a localização correspondentes ao IP
        ip_info = df[df['IP'] == ip_to_connect]

        if not ip_info.empty:
            codigo = ip_info['CODIGO'].values[0]
            localizacao = ip_info['LOCALIZACAO'].values[0]

            print(f'Conectando ao IP {ip_to_connect} (Código: {codigo}, Localização: {localizacao})')

            # Abra o programa Tz0Console.exe ou encontre a janela se já estiver aberta
            program_path = r'C:\Program Files (x86)\Trauma Zer0\Tz0\ModulesApp\Tz0Console.exe'
            tz0_console_window = find_or_open_tz0_console(program_path)

            # Verificar se a janela do Tz0Console está ativa
            if tz0_console_window.isActive:
                # Clique dentro do campo de IP do Tz0
                pyautogui.click(x=214, y=189, clicks=2)
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.press('backspace')

                # Digite o IP diretamente da lista
                pyautogui.typewrite(ip_to_connect)

                # Aguarda um curto período de tempo para garantir que o IP seja colado
                time.sleep(2)

                # Clica no botão Conectar Tz0
                pyautogui.moveTo(x=631, y=189)
                pyautogui.click(x=631, y=189)
            else:
                print("A janela do Tz0 Console não está ativa.")
        else:
            print(f"IP {ip_to_connect} não encontrado no arquivo.")
    else:
        print("Nenhum IP com 'Ping OK' encontrado no arquivo.")
# Caminho do arquivo Excel
excel_file_path = r'C:\24 - RPA\Atualizacao_PDV_Py\Banco\IPs_PDVs.xlsx'

# Função para verificar se a imagem do botão de conexão está presente
def verificarBotao(excel_file_path, imagem_path):
    try:
        # Tente localizar a imagem na tela
        posicao = pyautogui.locateOnScreen(imagem_path)

        # Verifique se a imagem (botão de conexão) foi encontrada
        if posicao is not None:
            resultado = "Conectado"
            def enviaArquivo():
                # Verificar a conexão
                resultado_conexao = verificarBotao(excel_file_path, imagem_path)
                
                # Verificar se a conexão foi bem-sucedida
                if resultado_conexao == "Conectado":
                    time.sleep(2.5)
                    pyautogui.click(x=52, y=222)
                    time.sleep(0.5)
                    pyautogui.click(x=174, y=111)
                    time.sleep(0.5)
                    pyautogui.click(x=206, y=165)
                    time.sleep(0.5)
                    pyautogui.click(x=241, y=165)
                    time.sleep(0.5)
                    pyautogui.click(x=470, y=466)
                else:
                    print("Falha na conexão. A ação não será realizada.")
        else:
            resultado = "Falha conexão"

        # Carregue o arquivo Excel em um DataFrame
        df = pd.read_excel(excel_file_path)

        # Adicione a nova coluna "Conexão" ao DataFrame com o resultado
        df['Conexão'] = resultado

        # Salve o DataFrame atualizado de volta no arquivo Excel
        df.to_excel(excel_file_path, index=False, engine='openpyxl')

        print(f'Teste realizado, {resultado}')

    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")

# Caminho do arquivo Excel
excel_file_path = r'C:\24 - RPA\Atualizacao_PDV_Py\Banco\IPs_PDVs.xlsx'

# Caminho da imagem do botão de conexão
imagem_path = r'C:\24 - RPA\Atualizacao_PDV_Py\Banco\botao_conexao.png'



