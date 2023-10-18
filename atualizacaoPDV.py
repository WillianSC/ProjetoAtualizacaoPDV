# Chamar as funções
from abreTrauma import abrirTrauma
from consultaBanco import conectar_banco
from relatorioIPs import consultar_e_salvar_em_excel
from testePing import testarPing, excel_file_path
from conectaIpOk import conectaIP

#Conecta no Banco e Salva o Relatório de IP das Lojas
connection, connection_string = conectar_banco()
consultar_e_salvar_em_excel(connection, connection_string)

# Realizar teste de Ping em todos os IPs contidos no relatorio de IP
# Chamada da função para realizar o teste de ping
testarPing(excel_file_path)

# Separar numa coluna os Pings com sucesso e falha - OK

#Abre o Tz0
abrirTrauma()

# Conectar IP contido no relatório ping "ok" e salvar num log a Loja e o PDV conectado
conectaIP(excel_file_path)

# Realiza um teste se a conexão foi bem sucedida. Caso sim, segue para envio. Caso não, guarda log e passa para proxima loja.


# Enviar o arquivo e após envio, desconectar e preparar para envio de um proximo PDV.
