import time
from abreTrauma import abrirTrauma
from consultaBanco import conectar_banco
from enviaArquivo import enviaArquivo
from relatorioIPs import consultar_e_salvar_em_excel
from testePing import testarPing, excel_file_path
from conectaIpOk import conectaIP
from validaConexaoRemota import verificarBotao, imagem_path

def atualizacaoPDV():
    # Conectar no Banco e Salvar o Relatório de IP das Lojas
    connection, connection_string = conectar_banco()
    consultar_e_salvar_em_excel(connection, connection_string)

    # Realizar teste de Ping em todos os IPs contidos no relatório de IP
    testarPing(excel_file_path)

    # Abre o Tz0
    abrirTrauma()

    # Conectar IP contido no relatório ping "ok" e salvar num log a Loja e o PDV conectado
    conectaIP(excel_file_path)

    # Realiza um teste se a conexão foi bem sucedida. Caso sim, segue para envio. Caso não, guarda log e passa para próxima loja.
    resultado_conexao = verificarBotao(excel_file_path, imagem_path)

    # Caso não encontre o botão de "desconectar", significa que a conexão com o trauma falhou, e ele deve passar para a próxima linha.

    # Enviar o arquivo e após envio, desconectar e preparar para envio de um próximo PDV.
    if resultado_conexao == "Conectado":
        enviaArquivo()
    else:
       print("Falha na conexão. A ação não será realizada.")

# Chamada da função principal
atualizacaoPDV()
