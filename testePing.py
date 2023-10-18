import pandas as pd
from ping3 import ping

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

# Chamada da função para realizar o teste de ping
testarPing(excel_file_path)
