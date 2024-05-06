import pandas as pd
from sqlalchemy import create_engine

server = '10.0.48.14'
database = 'millennium'
username = 'winsql'
password = 'winsql2020'

conn_str = f'mssql+pyodbc://{username}:{password}@{server}/{database}?driver=SQL+Server'

engine = create_engine(conn_str)

query = """SELECT NROEMPRESA,
	   plataforma ,
	   pedidoIfood ,
       Nro_SMR,
	   Data_Carga ,
	   Endereco ,
	   Periodo,
	   Horario_Inicial,
	   Municipio,
	   CEP,
	   Status_SMR,
	   Situacao_Pedido_C5,
	   Cliente,
	   Tipo_Frete as Tip_Frete
FROM NGM_PEDIDOS_ONDA_C5 npoc
WHERE CONVERT(DATE, npoc.Data_Carga) >= CONVERT(DATE, GETDATE())"""

dados_onda = pd.read_sql_query(query, engine)

caminho_arquivo_excel = r'C:\Users\erik.brigido\OneDrive - nagumo.com.br\ECOMMERCE\Logística\ONDA DE SEPARAÇÃO\dados_onda_atual.xlsx'

dados_onda.to_excel(caminho_arquivo_excel, index=False)
