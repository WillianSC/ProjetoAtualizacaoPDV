import time
import pyautogui
from validaConexaoRemota import verificarBotao, excel_file_path, imagem_path

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

# Chamada da função para enviar o arquivo, que verificará a conexão antes de executar a ação
# enviaArquivo()
