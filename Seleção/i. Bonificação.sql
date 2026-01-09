SELECT
    COD_CLIENTE,
    NOME_CLIENTE,
    -- Valor total das VENDAS
    SUM(CASE 
        WHEN DESC_TIPO_OPERACAO LIKE '%VENDA%' 
        THEN VALOR 
        ELSE 0 
    END) AS VALOR_VENDAS,
    
    -- Quantidade de VENDAS (contando registros de venda)
    COUNT(CASE 
        WHEN DESC_TIPO_OPERACAO LIKE '%VENDA%' 
        THEN 1 
    END) AS QTDE_VENDAS,
    
    -- Valor total das BONIFICAÇÕES
    SUM(CASE 
        WHEN DESC_TIPO_OPERACAO LIKE '%BONIFICACAO%' 
        THEN VALOR 
        ELSE 0 
    END) AS VALOR_BONIFICACOES,
    
    -- Percentual Bonificação (%):
    CASE
        WHEN SUM(CASE 
                    WHEN DESC_TIPO_OPERACAO LIKE '%VENDA%' 
                    THEN VALOR 
                    ELSE 0 
                 END) > 0
        THEN (
            SUM(CASE 
                    WHEN DESC_TIPO_OPERACAO LIKE '%BONIFICACAO%' 
                    THEN VALOR 
                    ELSE 0 
                END) 
            / 
            SUM(CASE 
                    WHEN DESC_TIPO_OPERACAO LIKE '%VENDA%' 
                    THEN VALOR 
                    ELSE 0 
                END) 
        )
        ELSE 0
    END AS PERCENTUAL_BONIFICACAO,
    
    -- Campos de grupo (primeiro valor encontrado por cliente)
    MAX(COD_ESTABELECIMENTO) AS COD_ESTABELECIMENTO,
    MAX(CASE
        WHEN COD_ESTABELECIMENTO IN ('R261', 'R221', 'R222', 'R541', 'R591', 'R281', 'R282', 'R283', 'R611', 'R121', 'R831', 'R351', 'R352', 'R461', 'R521') THEN 'AVANÇAR'
        WHEN COD_ESTABELECIMENTO IN ('R201', 'R311', 'R312', 'R313', 'R191', 'R781', 'R301') THEN 'BASE'
        WHEN COD_ESTABELECIMENTO IN ('R031', 'R041', 'R291', 'R292', 'R631', 'R641', 'R791', 'R801') THEN 'CRESCER'
        WHEN COD_ESTABELECIMENTO IN ('R651', 'R671', 'R681', 'R021', 'R181', 'R691', 'R131', 'R721', 'R751') THEN 'FORTALEZA'
        WHEN COD_ESTABELECIMENTO IN ('R211', 'R341', 'R451', 'R481', 'R711', 'R231', 'R234', 'R471', 'R472', 'R061', 'R531') THEN 'PLANALTO'
        WHEN COD_ESTABELECIMENTO IN ('R071', 'R074', 'R382', 'R501', 'R502', 'R661', 'R701', 'R491', 'R492', 'R241', 'R243', 'R621', 'R761', 'R371', 'R731', 'R821') THEN 'SUL'
        WHEN COD_ESTABELECIMENTO IN ('R011', 'R511', 'R101', 'R811', 'R051', 'R052', 'R161') THEN 'VISÃO'
        ELSE 'CADASTRAR'
    END) AS GRUPO
FROM 
    VW_AUDIT_RM_ORDENS_VENDA
WHERE 
    COD_ESTABELECIMENTO = 'R631'
    AND DATA_NOTA_FISCAL BETWEEN '2025-12-01' AND '2025-12-31' 
    AND PARA_FATURAMENTO = 'Sim'
    AND (DESC_TIPO_OPERACAO LIKE '%VENDA%' OR DESC_TIPO_OPERACAO LIKE '%BONIFICACAO%')
GROUP BY
    COD_CLIENTE,
    NOME_CLIENTE
ORDER BY
    COD_CLIENTE
