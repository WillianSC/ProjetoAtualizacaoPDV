import pandas as pd
import pygetwindow as gw
import subprocess
import time
import pyautogui
import time

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

            # Use o pygetwindow para encontrar a janela do Tz0Console
            tz0_console_window = gw.getWindowsWithTitle('Tz0 Console')[0]

            # Agora você pode implementar o código para copiar o IP e conectá-lo no Tz0Console
            # Clique em um ponto específico na tela
            #pyautogui.click(x=221, y=189)


        else:
            print(f"IP {ip_to_connect} não encontrado no arquivo.")
    else:
        print("Nenhum IP com 'Ping OK' encontrado no arquivo.")

# Caminho do arquivo Excel
excel_file_path = r'C:\24 - RPA\Atualizacao_PDV_Py\Banco\IPs_PDVs.xlsx'

# Chamada da função para conectar ao IP com "Ping OK"
conectaIP(excel_file_path)
