import pyautogui
import pandas as pd

# Função para verificar se a imagem do botão de conexão está presente
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

        print(f'Teste realizado, {resultado}')

    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")

# Caminho do arquivo Excel
excel_file_path = r'C:\24 - RPA\Atualizacao_PDV_Py\Banco\IPs_PDVs.xlsx'

# Caminho da imagem do botão de conexão
imagem_path = r'C:\24 - RPA\Atualizacao_PDV_Py\Banco\botao_conexao.png'

# Chamada da função para realizar o teste do botão e atualizar a planilha
verificarBotao(excel_file_path, imagem_path)