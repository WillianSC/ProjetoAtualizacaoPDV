import logging
import pandas as pd
import pygetwindow as gw
import subprocess
import time
import pyautogui

# Constante - Caminho do arquivo Excel
excel_file_path = r'C:\24 - RPA\Atualizacao_PDV_Py\Banco\IPs_PDVs.xlsx'

# Constante - Caminho da imagem do botão de conexão
imagem_path = r'C:\24 - RPA\Atualizacao_PDV_Py\Banco\botao_conexao.PNG'
imagem_envio_completo = r'C:\24 - RPA\Atualizacao_PDV_Py\Banco\botao_completo.PNG'

# Logging
logging.basicConfig(filename='rpa_log.txt', level=logging.INFO,
                    format='%(asctime)s | %(levelname)s | %(message)s', datefmt='%d-%m-%Y %H:%M:%S')

def conectaIP(excel_file_path):

    # Ler o arquivo Excel
    df = pd.read_excel(excel_file_path)

    # Filtrar os IPs com "Ping OK"
    df_ping_ok = df[df['Ping Result'] == 'Ping OK']

    # Obter a lista de IPs filtrados
    ips_ping_ok = df_ping_ok['IP'].tolist()

    # Loop através dos IPs com "Ping OK"
    for ip_to_connect in ips_ping_ok:
        time.sleep(3)
        # Abra o programa Tz0Console.exe para cada IP
        program_path = r'C:\Program Files (x86)\Trauma Zer0\Tz0\ModulesApp\Tz0Console.exe'
        process = subprocess.Popen(program_path)

        # Aguarde o programa carregar
        time.sleep(20)
        
        # Encontre o código e a localização correspondentes ao IP
        ip_info = df[df['IP'] == ip_to_connect]

        if not ip_info.empty:
            codigo = ip_info['CODIGO'].values[0]
            localizacao = ip_info['LOCALIZACAO'].values[0]

            print(f'Conectando ao IP {ip_to_connect} (Código: {codigo}, Localização: {localizacao})')
            logging.info(f'{time.strftime("%d-%m-%Y %H:%M")} | IP {ip_to_connect} - Código: {codigo} - Localização: {localizacao} - Tentando conexão')

            # Clica dentro do campo de IP do Tz0
            pyautogui.click(x=214, y=189, clicks = 2)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')

            # Digita o IP diretamente da lista
            pyautogui.typewrite(ip_to_connect)

            # Aguarda um curto período de tempo para garantir que o IP seja colado
            time.sleep(2)

            # Clica no botão Conectar Tz0
            pyautogui.moveTo(x=629, y=192)
            pyautogui.click(x=629, y=192)
            time.sleep(5)
            
            verificar_conexao = verificarBotao(excel_file_path, imagem_path)
            if verificar_conexao == "Conectado":
                enviaArquivo()
                process.terminate()
            else:
                process.terminate()
        else:
            print(f"IP {ip_to_connect} não encontrado no arquivo.")
            logging.info(f"IP {ip_to_connect} não encontrado no arquivo.")
            process.terminate()
    else:
        print("Nenhum IP com 'Ping OK' encontrado no arquivo.")
        logging.info("Nenhum IP com 'Ping OK' encontrado no arquivo.")
        process.terminate()
    
def verificarBotao(excel_file_path, imagem_path):
    time.sleep(4)
    try:
        x_esquerda, y_superior = 0, 633
        largura, altura = 175, 94
        area_busca = (x_esquerda, y_superior, largura, altura)

        # Capture a tela na área específica
        tela = pyautogui.screenshot(region=area_busca)

        # Tente localizar a imagem de referência na tela capturada
        posicao = pyautogui.locateOnScreen(imagem_path, region=area_busca)

        # Verifique se a posição da imagem foi encontrada
        if posicao is not None:
            resultado = "Conectado"
            logging.info(f'{time.strftime("%d-%m-%Y %H:%M")}|Botão desconectar encontrado - {resultado}')
            time.sleep(2)
        else:
            resultado = "Falha conexão"
            logging.warning(f'{time.strftime("%d-%m-%Y %H:%M")}|Falha ao verificar o botão de desconexão.')
            time.sleep(2)
            pass

        print(f'Botão desconectar verificado, {resultado}')
        logging.info(f'Botão desconectar verificado, {resultado}')
        return resultado  # Retorna o resultado da verificação

    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        logging.error(f'{time.strftime("%d-%m-%Y %H:%M")}| Ocorreu um erro: {str(e)}')
        return "Erro"  # Retorna "Erro" em caso de exceção


def enviaArquivo():
    tentativa_atual = 0
    tempo_limite = 40  # Tempo limite total em segundos
    intervalo_tentativa = 2  # Intervalo entre tentativas em segundos

    # Abre o menu "UPLOAD/NORMAL"
    time.sleep(10)
    pyautogui.click(x=52, y=222)
    time.sleep(0.5)
    pyautogui.click(x=174, y=111)
    time.sleep(0.5)
    pyautogui.click(x=206, y=165)
    time.sleep(0.5)
    pyautogui.click(x=241, y=165)
    time.sleep(0.5)
    pyautogui.click(x=470, y=466)

    # Vai até o caminho onde se encontra o arquivo ZIP a ser enviado
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('space')
    time.sleep(0.5)
    pyautogui.press('backspace')
    time.sleep(1)
    pyautogui.write(r"C:\Atualizacao", interval=0.01)
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.doubleClick(x=954, y=362)
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

    while tentativa_atual * intervalo_tentativa < tempo_limite:
        checarEnvio = pyautogui.locateOnScreen(imagem_envio_completo)

        if checarEnvio is not None:
            resultadoEnvio = "Enviado"
            logging.info(f'{time.strftime("%d-%m-%Y %H:%M")}|Envio identificado')
            print("Envio identificado")
            time.sleep(2)
            break  # Saia do loop se a imagem for encontrada
        else:
            resultadoEnvio = "Aguardando envio"
            logging.info(f'{time.strftime("%d-%m-%Y %H:%M")}|Imagem de envio completo não encontrado na tentativa {tentativa_atual}')
            print('|Imagem de envio completo não encontrado na tentativa')
            tentativa_atual += 1
            time.sleep(intervalo_tentativa)  # Aguarde pelo intervalo antes da próxima tentativa

    print({resultadoEnvio})
    logging.info({resultadoEnvio})

# Chamada da função para conectar ao IP com "Ping OK"
conectaIP(excel_file_path)