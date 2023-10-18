import time
import pyautogui
import pygetwindow as gw
import subprocess

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

# COMANDO PARA MANTER O CODIGO EM EXECUÇÃO PARA TESTE#
#input("Pressione Enter para encerrar o programa.")  #
# COMANDO PARA MANTER O CODIGO EM EXECUÇÃO PARA TESTE#