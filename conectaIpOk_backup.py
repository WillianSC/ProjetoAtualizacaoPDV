import pandas as pd
import pygetwindow as gw
import subprocess
import time
import pyautogui

# Constante - Caminho do arquivo Excel
excel_file_path = r'C:\24 - RPA\Atualizacao_PDV_Py\Banco\IPs_PDVs.xlsx'

# Constante - Caminho da imagem do botão de conexão
imagem_path = r'C:\24 - RPA\Atualizacao_PDV_Py\Banco\botao_conexao.png'
imagem_envio = r'C:\24 - RPA\Atualizacao_PDV_Py\Banco\botao_completo.PNG'

def conectaIP(excel_file_path):
    # Ler o arquivo Excel
    df = pd.read_excel(excel_file_path)

    # Filtrar os IPs com "Ping OK"
    df_ping_ok = df[df['Ping Result'] == 'Ping OK']

    # Obter a lista de IPs filtrados
    ips_ping_ok = df_ping_ok['IP'].tolist()

    # Loop através dos IPs com "Ping OK"
    for ip_to_connect in ips_ping_ok:
        # Encontre o código e a localização correspondentes ao IP
        ip_info = df[df['IP'] == ip_to_connect]

        if not ip_info.empty:
            codigo = ip_info['CODIGO'].values[0]
            localizacao = ip_info['LOCALIZACAO'].values[0]

            print(f'Conectando ao IP {ip_to_connect} (Código: {codigo}, Localização: {localizacao})')

            # Abra o programa Tz0Console.exe
            program_path = r'C:\Program Files (x86)\Trauma Zer0\Tz0\ModulesApp\Tz0Console.exe'
            process = subprocess.Popen(program_path)

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
            pyautogui.moveTo(x=629, y=192)
            pyautogui.click(x=629, y=192)
            time.sleep(5)

            enviaArquivo()

            # Aguarda a detecção da imagem de envio antes de prosseguir
            resultado_envio = aguardaEnvio(imagem_envio, df)

            if resultado_envio == "Enviado":
                # A imagem foi detectada, então prossegue com o envio do arquivo
                process.terminate()
            elif resultado_envio == "Aguardando envio":
                # A imagem não foi detectada mesmo após a espera, trata como uma falha
                print("Falha no envio. Imagem não detectada após o tempo de espera.")
                process.terminate()
            else:
                # Outro cenário de falha
                print("Falha no envio. Consulte o resultado_envio para mais informações.")
                process.terminate()
        else:
            print(f"IP {ip_to_connect} não encontrado no arquivo.")
    else:
        print("Nenhum IP com 'Ping OK' encontrado no arquivo.")

def aguardaEnvio(imagem_envio, df):
    max_tentativas = 10  # Número máximo de tentativas
    tentativa_atual = 0

    try:
        while tentativa_atual < max_tentativas:
            # Tente localizar a imagem na tela
            posicaoenvio = pyautogui.locateOnScreen(imagem_envio)

            # Verifique se a imagem (botão de envio) foi encontrada
            if posicaoenvio is not None:
                resultadoenvio = "Enviado"
                break  # Saia do loop se a imagem for encontrada
            else:
                resultadoenvio = "Aguardando envio"
                tentativa_atual += 1
                time.sleep(1)  # Aguarde por 3 segundos antes da próxima tentativa

        # Se a imagem não for encontrada após o número máximo de tentativas
        if tentativa_atual == max_tentativas:
            resultadoenvio = "Falha envio"

            # Adicione a nova coluna "Conexão" ao DataFrame com o resultado
            df['StatusEnvio'] = resultadoenvio

            # Salve o DataFrame atualizado de volta no arquivo Excel
            df.to_excel(excel_file_path, index=False, engine='openpyxl')

            print(f'Teste realizado, {resultadoenvio}')

        return resultadoenvio

    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        return "Erro"  # Retorna "Erro" em caso de exceção
    
def verificarBotao(excel_file_path, imagem_path):
    try:
        # Tente localizar a imagem na tela
        posicao = pyautogui.locateOnScreen(imagem_path)

        # Verifique se a imagem (botão de conexão) foi encontrada
        if posicao is not None:
            resultado = "Conectado"
        else:
            resultado = "Falha conexão"

        # Carregue o arquivo Excel em um DataFrame
        df = pd.read_excel(excel_file_path)

        # Adicione a nova coluna "Conexão" ao DataFrame com o resultado
        df['Conexão'] = resultado

        # Salve o DataFrame atualizado de volta no arquivo Excel
        df.to_excel(excel_file_path, index=False, engine='openpyxl')

        print(f'Botão desconectar verificado, {resultado}')

        return resultado  # Retorna o resultado da verificação
    

    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")
        return "Erro"  # Retorna "Erro" em caso de exceção

def enviaArquivo():
    # Verificar a conexão e atribuir o resultado a resultado_conexao
    resultado_conexao = verificarBotao(excel_file_path, imagem_path)

    # Verificar se a conexão foi bem-sucedida
    if resultado_conexao == "Conectado":
        #Abre o menu "UPLOAD/NORMAL"
        time.sleep(20)
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
    else:
        print("Falha na conexão. A ação não será realizada.")
        pass

# Chamada da função para conectar ao IP com "Ping OK"
conectaIP(excel_file_path)