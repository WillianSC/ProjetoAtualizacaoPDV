import pandas as pd
import pygetwindow as gw
import subprocess
import time
import pyautogui

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

            # Abra o programa Tz0Console.exe
            program_path = r'C:\Program Files (x86)\Trauma Zer0\Tz0\ModulesApp\Tz0Console.exe'
            subprocess.Popen(program_path)

            # Aguarde o programa carregar
            time.sleep(20)

            # Usar o pygetwindow para encontrar a janela do Tz0Console
            tz0_console_window = gw.getWindowsWithTitle('Tz0 Console')[0]

            # Clica dentro do campo de IP do Tz0
            pyautogui.click(x=214, y=189, clicks = 2)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')

            # Digita o IP diretamente da lista
            pyautogui.typewrite(ip_to_connect)

            # Aguarda um curto período de tempo para garantir que o IP seja colado
            time.sleep(2)

            # Clica no botão Conectar Tz0
            pyautogui.moveTo(x=631, y=189)
            pyautogui.click(x=631, y=189)

        else:
            print(f"IP {ip_to_connect} não encontrado no arquivo.")
    else:
        print("Nenhum IP com 'Ping OK' encontrado no arquivo.")

# Caminho do arquivo Excel
excel_file_path = r'C:\24 - RPA\Atualizacao_PDV_Py\Banco\IPs_PDVs.xlsx'

# Chamada da função para conectar ao IP com "Ping OK"
conectaIP(excel_file_path)

# COMANDO PARA MANTER O CODIGO EM EXECUÇÃO PARA TESTE#
input("Pressione Enter para encerrar o programa.")  #
# COMANDO PARA MANTER O CODIGO EM EXECUÇÃO PARA TESTE#
