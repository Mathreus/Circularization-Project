import pandas as pd
import pyodbc
import os
from datetime import datetime
from warnings import filterwarnings

# Filtrar o warning do pandas
filterwarnings('ignore', message='pandas only supports SQLAlchemy connectable')

# Configurações de conexão com o banco de dados
server = 'DCMDWF01A.MOURA.INT'  
database = 'ax'   
username = 'uAuditoria' 
password = '@ud!t0$!@202&22'  

# Definir o caminho de salvamento conforme solicitado
caminho_salvamento = r'C:\Users\matheus.melo\OneDrive - Acumuladores Moura SA\Documentos\Drive - Matheus Melo\Auditoria\2026\04. Circularização\Validações\Fluminense - R121'
nome_arquivo = f'analise_devolucoes_faturamento_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx'
caminho_completo = os.path.join(caminho_salvamento, nome_arquivo)

def conectar_banco():
    """Estabelece conexão com o banco de dados"""
    try:
        # String de conexão
        conn_str = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        
        # Estabelecer conexão
        conexao = pyodbc.connect(conn_str)
        print("Conexão estabelecida com sucesso!")
        return conexao
        
    except Exception as e:
        print(f"Erro ao conectar ao banco de dados: {e}")
        return None

def executar_query(conn, query):
    """Executa uma query e retorna DataFrame"""
    try:
        # Usar cursor.execute() em vez de pd.read_sql para evitar warnings
        cursor = conn.cursor()
        cursor.execute(query)
        
        # Obter nomes das colunas
        columns = [column[0] for column in cursor.description]
        
        # Buscar todos os resultados
        rows = cursor.fetchall()
        
        # Criar DataFrame
        df = pd.DataFrame.from_records(rows, columns=columns)
        
        return df
    except Exception as e:
        print(f"Erro ao executar query: {e}")
        return pd.DataFrame()

def executar_query_devolucoes(conn):
    """Executa a query de devoluções por grupo"""
    # Lista de estabelecimentos para AVANÇAR
    estabelecimentos_avancar = ['R261', 'R221', 'R222', 'R541', 'R591', 'R281', 'R282', 'R283', 'R611', 'R121', 'R831', 'R351', 'R352', 'R461', 'R521']
    
    query_devolucoes = f"""
    SELECT
        CASE
            WHEN COD_ESTABELECIMENTO IN ({','.join([f"'{e}'" for e in estabelecimentos_avancar])}) THEN 'AVANÇAR'
            ELSE 'VERIFICAR'
        END AS GRUPO_RM,
        SUM(QUANTIDADE) AS QUANTIDADE_DEVOLVIDO,
        SUM(VALOR) AS VALOR_DEVOLVIDO
    FROM 
        VW_AUDIT_RM_ORDENS_VENDA
    WHERE
        COD_ESTABELECIMENTO IN ({','.join([f"'{e}'" for e in estabelecimentos_avancar])})
        AND DATA_NOTA_FISCAL BETWEEN '2025-07-01' AND '2025-12-31' 
        AND PARA_FATURAMENTO = 'SIM'
        AND CFOP IN ('1.201', '1.202', '1.203', '1.204', '1.410', '1.411', '1.553', '1.660', '1.661', '1.662', 
                    '2.201', '2.202', '2.203', '2.204', '2.410', '2.411', '2.553', '2.660', '2.661', '2.662',
                    '3.201', '3.202', '3.211', '3.553')
    GROUP BY
        CASE
            WHEN COD_ESTABELECIMENTO IN ({','.join([f"'{e}'" for e in estabelecimentos_avancar])}) THEN 'AVANÇAR'
            ELSE 'VERIFICAR'
        END
    """
    
    try:
        df_devolucoes = executar_query(conn, query_devolucoes)
        print(f"Query de devoluções executada: {len(df_devolucoes)} registros encontrados")
        return df_devolucoes
    except Exception as e:
        print(f"Erro ao executar query de devoluções: {e}")
        return pd.DataFrame()

def executar_query_faturamento(conn):
    """Executa a query de faturamento por grupo"""
    # Lista de estabelecimentos para AVANÇAR
    estabelecimentos_avancar = ['R261', 'R221', 'R222', 'R541', 'R591', 'R281', 'R282', 'R283', 'R611', 'R121', 'R831', 'R351', 'R352', 'R461', 'R521']
    
    query_faturamento = f"""
    SELECT
        CASE
            WHEN COD_ESTABELECIMENTO IN ({','.join([f"'{e}'" for e in estabelecimentos_avancar])}) THEN 'AVANÇAR'
            ELSE 'VERIFICAR'
        END AS GRUPO_RM,
        SUM(QUANTIDADE) AS QUANTIDADE_VENDAS,
        SUM(VALOR) AS VALOR_VENDAS
    FROM 
        VW_AUDIT_RM_ORDENS_VENDA
    WHERE 
        COD_ESTABELECIMENTO IN ({','.join([f"'{e}'" for e in estabelecimentos_avancar])})
        AND DATA_NOTA_FISCAL BETWEEN '2025-07-01' AND '2025-12-31'  
        AND PARA_FATURAMENTO = 'SIM'
        AND CFOP IN ('5.100', '5.101', '5.102', '5.103', '5.104', '5.105', '5.106', '5.109', '5.110', '5.111', 
                    '5.112', '5.113', '5.114', '5.115', '5.116', '5.117', '5.118', '5.119', '5.120', '5.122', 
                    '5.123', '5.250','5.251', '5.252', '5.253', '5.254', '5.255', '5.256', '5.257', '5.258', 
                    '5.401', '5.402', '5.403', '5.405', '5.651', '5.652', '5.653', '5.654', '5.655', '5.656',
                    '5.667', '6.101', '6.102', '6.103','6.104', '6.105', '6.106', '6.107', '6.108', '6.109',
                    '6.110', '6.111', '6.112', '6.113', '6.114', '6.115', '6.116', '6.117', '6.118', '6.119',
                    '6.120', '6.122', '6.123', '6.250', '6.251', '6.252', '6.253', '6.254', '6.255', '6.256',
                    '6.257', '6.258', '6.401', '6.402', '6.403', '6.404', '6.651', '6.652', '6.653', '6.654',
                    '6.655', '6.656', '6.667', '7.100', '7.101', '7.102', '7.105', '7.106','7.127', '7.250', 
                    '7.251', '7.651', '7.654', '7.667')
    GROUP BY
        CASE
            WHEN COD_ESTABELECIMENTO IN ({','.join([f"'{e}'" for e in estabelecimentos_avancar])}) THEN 'AVANÇAR'
            ELSE 'VERIFICAR'
        END
    """
    
    try:
        df_faturamento = executar_query(conn, query_faturamento)
        print(f"Query de faturamento executada: {len(df_faturamento)} registros encontrados")
        return df_faturamento
    except Exception as e:
        print(f"Erro ao executar query de faturamento: {e}")
        return pd.DataFrame()

def formatar_numeros(df):
    """Formata todas as colunas numéricas com 2 casas decimais"""
    # Identificar colunas numéricas
    colunas_numericas = df.select_dtypes(include=['float64', 'int64']).columns
    
    for coluna in colunas_numericas:
        if coluna == 'QUANTIDADE_DEVOLVIDO' or coluna == 'QUANTIDADE_VENDAS':
            # Para quantidades, formatar como inteiro ou com 0 casas decimais
            df[coluna] = df[coluna].apply(lambda x: f"{x:,.0f}" if pd.notnull(x) else "0")
        else:
            # Para valores monetários, formatar com 2 casas decimais
            df[coluna] = df[coluna].apply(lambda x: f"{x:,.2f}" if pd.notnull(x) else "0.00")
    
    return df

def calcular_taxa_devolucao(df_devolucoes, df_faturamento):
    """Calcula a taxa de devolução (devolução/faturamento) por grupo"""
    
    # Verificar se temos dados
    if df_devolucoes.empty or df_faturamento.empty:
        print("AVISO: Uma das queries não retornou dados.")
        
        # Criar DataFrames vazios com estrutura correta
        df_resultado = pd.DataFrame({
            'GRUPO_RM': ['AVANÇAR', 'VERIFICAR'],
            'QUANTIDADE_DEVOLVIDO': [0.0, 0.0],
            'VALOR_DEVOLVIDO': [0.0, 0.0],
            'QUANTIDADE_VENDAS': [0.0, 0.0],
            'VALOR_VENDAS': [0.0, 0.0]
        })
    else:
        # Realizar merge das duas tabelas usando GRUPO_RM como chave
        df_merge = pd.merge(df_devolucoes, df_faturamento, 
                            on=['GRUPO_RM'], 
                            how='outer', 
                            suffixes=('_DEV', '_FAT'))
        
        # Preencher valores nulos com 0 para grupos sem dados
        for col in ['QUANTIDADE_DEVOLVIDO', 'VALOR_DEVOLVIDO', 'QUANTIDADE_VENDAS', 'VALOR_VENDAS']:
            if col in df_merge.columns:
                df_merge[col] = df_merge[col].fillna(0)
            else:
                df_merge[col] = 0.0
        
        df_resultado = df_merge
    
    # Calcular taxa de devolução (VALOR_DEVOLVIDO / VALOR_VENDAS)
    def calcular_taxa(valor_dev, valor_fat):
        if valor_fat == 0:
            return 0.0
        return (valor_dev / valor_fat)
    
    df_resultado['TAXA_DEVOLUCAO_VALOR'] = df_resultado.apply(
        lambda x: calcular_taxa(x['VALOR_DEVOLVIDO'], x['VALOR_VENDAS']), 
        axis=1
    )
    
    # Calcular taxa de devolução por quantidade
    df_resultado['TAXA_DEVOLUCAO_QUANTIDADE'] = df_resultado.apply(
        lambda x: calcular_taxa(x['QUANTIDADE_DEVOLVIDO'], x['QUANTIDADE_VENDAS']), 
        axis=1
    )
    
    # Formatar as taxas como porcentagem com 2 casas decimais
    df_resultado['TAXA_DEVOLUCAO_VALOR_PCT'] = df_resultado['TAXA_DEVOLUCAO_VALOR'].apply(
        lambda x: f"{x:.2%}"
    )
    
    df_resultado['TAXA_DEVOLUCAO_QUANTIDADE_PCT'] = df_resultado['TAXA_DEVOLUCAO_QUANTIDADE'].apply(
        lambda x: f"{x:.2%}"
    )
    
    # Criar cópia para formatação de exibição
    df_formatado = df_resultado.copy()
    
    # Formatar colunas numéricas
    df_formatado = formatar_numeros(df_formatado)
    
    # Manter as taxas como números para cálculos
    df_resultado['TAXA_DEVOLUCAO_VALOR_FORMATADA'] = df_resultado['TAXA_DEVOLUCAO_VALOR'].apply(lambda x: f"{x:.4f}")
    df_resultado['TAXA_DEVOLUCAO_QUANTIDADE_FORMATADA'] = df_resultado['TAXA_DEVOLUCAO_QUANTIDADE'].apply(lambda x: f"{x:.4f}")
    
    # Ordenar por grupo
    df_resultado = df_resultado.sort_values('GRUPO_RM')
    df_formatado = df_formatado.sort_values('GRUPO_RM')
    
    return df_resultado, df_formatado

def salvar_para_excel(df_resultado, df_formatado, caminho_completo):
    """Salva os resultados em um arquivo Excel no caminho especificado"""
    
    try:
        # Verificar se o diretório existe, se não, criar
        diretorio = os.path.dirname(caminho_completo)
        if not os.path.exists(diretorio):
            os.makedirs(diretorio, exist_ok=True)
            print(f"Diretório criado: {diretorio}")
        
        # Criar um ExcelWriter para salvar múltiplas abas
        with pd.ExcelWriter(caminho_completo, engine='openpyxl') as writer:
            # Salvar análise por grupo formatada
            df_formatado.to_excel(writer, sheet_name='Analise_por_Grupo', index=False)
            
            # Salvar dados brutos para referência
            df_bruto = df_resultado.copy()
            df_bruto.to_excel(writer, sheet_name='Dados_Brutos', index=False)
            
            # Adicionar uma aba com informações sobre a análise
            info_df = pd.DataFrame({
                'Campo': ['Data de Execução', 'Período Analisado', 'Estabelecimento Foco', 'Total de Grupos'],
                'Valor': [
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    '2025-07-01 a 2025-12-31',
                    'R121 (Fluminense)',
                    len(df_resultado)
                ]
            })
            info_df.to_excel(writer, sheet_name='Informacoes', index=False)
        
        print(f"Arquivo Excel salvo com sucesso em: {caminho_completo}")
        return True
        
    except Exception as e:
        print(f"Erro ao salvar arquivo Excel: {e}")
        print("Tentando salvar no diretório atual como fallback...")
        
        # Tentativa de fallback: salvar no diretório atual
        try:
            caminho_fallback = f'analise_devolucoes_faturamento_{datetime.now().strftime("%Y%m%d_%H%M")}_fallback.xlsx'
            df_formatado.to_excel(caminho_fallback, index=False)
            
            print(f"Arquivo salvo no diretório atual como: {caminho_fallback}")
            print(f"Caminho atual: {os.getcwd()}")
            return True
            
        except Exception as e2:
            print(f"Erro no fallback: {e2}")
            return False

def main():
    """Função principal"""
    
    print(f"Destino do arquivo: {caminho_completo}")
    
    # Conectar ao banco de dados
    conn = conectar_banco()
    if not conn:
        print("Não foi possível conectar ao banco de dados. Verifique as credenciais.")
        return
    
    try:
        # Executar queries
        df_devolucoes = executar_query_devolucoes(conn)
        df_faturamento = executar_query_faturamento(conn)
        
        # Exibir preview dos dados brutos
        if not df_devolucoes.empty:
            print("\nPrévia dos dados de devoluções (brutos):")
            print(df_devolucoes.to_string(index=False))
        
        if not df_faturamento.empty:
            print("\nPrévia dos dados de faturamento (brutos):")
            print(df_faturamento.to_string(index=False))
        
        # Calcular taxa de devolução
        df_resultado, df_formatado = calcular_taxa_devolucao(df_devolucoes, df_faturamento)
        
        # Salvar resultados em arquivo Excel
        sucesso = salvar_para_excel(df_resultado, df_formatado, caminho_completo)
        
        if sucesso:
            print("\n" + "="*60)
            print("ANÁLISE DE DEVOLUÇÃO POR FATURAMENTO - R121 FLUMINENSE")
            print("="*60)
            
            # Exibir resultados formatados
            print("\nRESULTADOS POR GRUPO (FORMATADOS):")
            print(df_formatado.to_string(index=False))
            
            # Exibir estatísticas gerais com formatação
            total_devolucao = df_resultado['VALOR_DEVOLVIDO'].sum()
            total_faturamento = df_resultado['VALOR_VENDAS'].sum()
            taxa_geral = total_devolucao / total_faturamento if total_faturamento > 0 else 0
            
            print(f"\nESTATÍSTICAS GERAIS:")
            print(f"Total de Devolução: R$ {total_devolucao:,.2f}")
            print(f"Total de Faturamento: R$ {total_faturamento:,.2f}")
            print(f"Taxa Geral de Devolução: {taxa_geral:.2%}")
            print(f"Total de grupos analisados: {len(df_resultado)}")
            
            # Exibir detalhes dos cálculos
            print(f"\nDETALHES DOS CÁLCULOS:")
            for _, row in df_resultado.iterrows():
                print(f"\nGrupo: {row['GRUPO_RM']}")
                print(f"  Devolução: R$ {row['VALOR_DEVOLVIDO']:,.2f} / Faturamento: R$ {row['VALOR_VENDAS']:,.2f}")
                print(f"  Taxa calculada: {row['TAXA_DEVOLUCAO_VALOR_FORMATADA']} = {row['TAXA_DEVOLUCAO_VALOR_PCT']}")
            
        else:
            print("\nATENÇÃO: Não foi possível salvar o arquivo no caminho especificado.")
            print("Verifique as permissões do diretório ou se o caminho está correto.")
        
    except Exception as e:
        print(f"Erro durante a execução: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if conn:
            conn.close()
            print("\nConexão com o banco de dados fechada.")

if __name__ == "__main__":
    main()
