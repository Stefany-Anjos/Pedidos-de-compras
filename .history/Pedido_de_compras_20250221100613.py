import psycopg2
from selenium.common.exceptions import TimeoutException
import datetime
from datetime import datetime

# Pega apenas a data atual (sem hora)
data_atual = datetime.now().date()
# Formata a data para o padrão dia, mês e ano
data_atual = f"{data_atual.strftime('%d-%m-%y')}"
print(f"Data atual: {data_atual}")

# Obtem o caminho da planilha
caminho_planilha = "M:\TI\ROBOS\ROBOS_PRONTOS\SEMINOVOS\Robo Pedido de Compras - OK\Planilhas"

# conexaão com o banco de dados
def conexaoBD():
   dbname = 'dukarsys'
   user = 'postgres'
   senha = 'S@ntos67'
   host = '192.168.20.5'
   port = '5432'

   try:

    # configurações para o banco de dados
    conn = psycopg2.connect(
        dbname=dbname,
        user=user,
        password=senha,
        host=host,
        port=port,
        options="-c client_encoding=UTF8"
    )

    # Criando um cursor
    cursor = conn.cursor()
   except TimeoutException:
         print("Erro na conexão")
def cria_plamnilha():
    nome_planilha = f''