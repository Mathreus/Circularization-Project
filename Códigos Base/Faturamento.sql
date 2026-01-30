# Bibliotecas base de conexão:
import pyodbc
import pandas as pd
from datetime import datetime
import os

# Defina as informações de conexão
server = 'DCMDWF01A.MOURA.INT'
database = 'ax'
username = 'uAuditoria'
password = '@ud!t0$!@202&22'
driver = 'SQL Server'  # Driver específico para o banco de dados que você está usando

# Construa a string de conexão
connection_string = f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}'

# Execute a consulta e salve em Excel
try:
    # Conecte-se ao banco de dados
    conexao = pyodbc.connect(connection_string)
    
    # Consulta SQL
    query = """
    SELECT
	    COD_ESTABELECIMENTO,
	    COD_CLIENTE,
	    NOME_CLIENTE,
	    UF_CLIENTE,
	    [CPF/CNPJ],
	    EMISSOR,
	    DATA_NOTA_FISCAL,
	    NUM_NOTA_FISCAL,
	    COD_ITEM,
	    DESCR_ITEM,
	    QUANTIDADE,
	    PESO_BIN,
	    VALOR,
	    CFOP,
	    DESC_TIPO_OPERACAO,
	    CASE
    	    WHEN COD_ESTABELECIMENTO IN ('R261', 'R221', 'R222', 'R541', 'R591', 'R281', 'R282', 'R283', 'R611', 'R121', 'R831', 'R351', 'R352', 'R461', 'R521') THEN 'AVANÇAR'
		    WHEN COD_ESTABELECIMENTO IN ('R201', 'R311', 'R312', 'R313', 'R191', 'R781', 'R301') THEN 'BASE'
		    WHEN COD_ESTABELECIMENTO IN ('R031', 'R041', 'R291', 'R292', 'R641', 'R791', 'R801') THEN 'CRESCER'
		    WHEN COD_ESTABELECIMENTO IN ('R651', 'R671', 'R681', 'R021', 'R181', 'R691', 'R131', 'R721', 'R751') THEN 'FORTALEZA'
		    WHEN COD_ESTABELECIMENTO IN ('R211', 'R341', 'R451', 'R481', 'R711', 'R231', 'R234', 'R471', 'R472', 'R061', 'R531') THEN 'PLANALTO'
		    WHEN COD_ESTABELECIMENTO IN ('R071', 'R074', 'R382', 'R501', 'R502', 'R661', 'R701', 'R491', 'R492', 'R241', 'R243', 'R621', 'R761', 'R371', 'R373', 'R731', 'R821') THEN 'SUL'
		    WHEN COD_ESTABELECIMENTO IN ('R011', 'R511', 'R101', 'R811', 'R051', 'R052', 'R161') THEN 'VISÃO'
    	    ELSE 'CADASTRAR'
	    END AS GRUPO,
	    LEN([CPF/CNPJ]) AS CARACTERES,
    	CASE
    	    WHEN LEN([CPF/CNPJ]) = 11 THEN 'FÍSICA'
    	    WHEN LEN([CPF/CNPJ]) = 14 THEN 'JURÍDICA'
    	    ELSE 'VERIFICAR'
	    END AS PESSOA,
	    VENDEDOR
    FROM 
	    VW_AUDIT_RM_ORDENS_VENDA
    WHERE 
	    COD_ESTABELECIMENTO = 'R121' -- Altere para o código da empresa que será circularizado 
	    AND DATA_NOTA_FISCAL BETWEEN '2025-07-07' AND '2026-01-07'  -- Altere a data base, lembrando de manter o formato ANO-MÊS-DIA
	    AND PARA_FATURAMENTO = 'Sim'
	    AND CFOP IN ('5.100', '5.101', '5.102', '5.103', '5.104', '5.105', '5.106', '5.109', '5.110', '5.111', '5.112', '5.113', '5.114', '5.115', '5.116', '5.117', '5.118', '5.119', '5.120', '5.122', '5.123', '5.250',
				    '5.251', '5.252', '5.253', '5.254', '5.255', '5.256', '5.257', '5.258', '5.401', '5.402', '5.403', '5.405', '5.651', '5.652', '5.653', '5.654', '5.655', '5.656', '5.667', '6.101', '6.102', '6.103',
				    '6.104', '6.105', '6.106', '6.107', '6.108', '6.109', '6.110', '6.111', '6.112', '6.113', '6.114', '6.115', '6.116', '6.117', '6.118', '6.119', '6.120', '6.122', '6.123', '6.250', '6.251', '6.252', 
				    '6.253', '6.254', '6.255', '6.256', '6.257', '6.258', '6.401', '6.402', '6.403', '6.404', '6.651', '6.652', '6.653', '6.654', '6.655', '6.656', '6.667', '7.100', '7.101', '7.102', '7.105', '7.106',
				    '7.127', '7.250', '7.251', '7.651', '7.654', '7.667')
        
    """
    
    # Executar a consulta diretamente com pandas para facilitar
    df = pd.read_sql_query(query, conexao)
    
    # Fechar a conexão
    conexao.close()
    
    # Verificar se há dados
    if len(df) > 0:
        # Definir o caminho para salvar o arquivo
        caminho_base = r'C:\Users\matheus.melo\OneDrive - Acumuladores Moura SA\Documentos\Drive - Matheus Melo\Auditoria\2026\04. Circularização\Validações\Fluminense - R121\Python'
        
        # Criar o diretório se não existir
        os.makedirs(caminho_base, exist_ok=True)
        
        # Nome do arquivo fixo como solicitado
        nome_arquivo = 'Base - Faturamento.xlsx'
        caminho_completo = os.path.join(caminho_base, nome_arquivo)
        
        # Salvar em Excel
        df.to_excel(caminho_completo, index=False, engine='openpyxl')
        
        print(f"✅ Arquivo salvo com sucesso!")
        print(f"📊 Total de registros: {len(df)}")
        print(f"📂 Caminho: {caminho_completo}")
        
        # Mostrar prévia dos dados
        print("\n📋 Prévia dos dados:")
        print(df.head())
        
    else:
        print("⚠️  Nenhum dado encontrado com os critérios especificados.")
        
except pyodbc.Error as e:
    print(f"❌ Erro na conexão ou consulta: {e}")
except Exception as e:
    print(f"❌ Erro inesperado: {e}")
