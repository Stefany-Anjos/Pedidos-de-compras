import psycopg2
from selenium.common.exceptions import TimeoutException


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