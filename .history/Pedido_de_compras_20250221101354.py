import os
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
caminho_planilha = r"M:\TI\ROBOS\ROBOS_PRONTOS\SEMINOVOS\Robo Pedido de Compras - OK\Planilhas"

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
    
    # Configura a query para buscar os pedidos de compra do banco de dados
    query_pedidoComp = """ select o.placa, o.renavan, s.descricao, c.nome_cliente
        from os o
        left join manutencao m on (m.OS=o.numero_os)
        left join servicos s on s.codigo_servico = o.cod_servico
        left join cliente c on(c.codigo_cliente=o.cod_cliente)
        where m.data_envio <> '30/12/1899'
        and m.data_retorno = '30/12/1899'
        and o.estado <>'C'
        and s.setor = 'S'
        and s.codigo_servico in (47,65)
        order by s.codigo_servico asc """
        
    # Executa a consulta de CPs para aprovar
    cursor.execute(query_pedidoComp) 
     
   except Exception as e:
         print(f"Erro na conexão: {e}")
# Cria planilha e renomeia         
def cria_plamnilha():
    nome_planilha = f'Pedido de compras {data_atual}.xlsx'
    caminho_completo = os.path.join(caminho_planilha, nome_planilha)
    
conexaoBD()
      