SELECT
    GRUPO,
    COD_ESTABELECIMENTO,
    PESSOA,
    COUNT(*) AS TOTAL_NOTAS,
    SUM(QUANTIDADE) AS TOTAL_QUANTIDADE,
    SUM(VALOR) AS TOTAL_VALOR,

    -- Métricas para devoluções:
    SUM(CASE WHEN TIPO_OPERACAO = 'DEVOLUCAO' THEN QUANTIDADE ELSE 0 END) AS QTD_DEVOLUCOES,
    
    -- Métricas Devoluções NF Própria:
    SUM(CASE WHEN TIPO_OPERACAO = 'DEVOLUCAO' AND EMISSOR = 'OwnEstablishment' THEN QUANTIDADE ELSE 0 END) AS QTD_DEVOLUCOES_NF_PROPRIA,

    -- Cálculo de índice de devolução por NF Própria:
    CASE 
        WHEN SUM(CASE WHEN TIPO_OPERACAO = 'DEVOLUCAO' THEN QUANTIDADE ELSE 0 END) > 0 
        THEN ROUND(
            CAST(SUM(CASE WHEN TIPO_OPERACAO = 'DEVOLUCAO' AND EMISSOR = 'OwnEstablishment' THEN QUANTIDADE ELSE 0 END) AS FLOAT) / 
            CAST(SUM(CASE WHEN TIPO_OPERACAO = 'DEVOLUCAO' THEN QUANTIDADE ELSE 0 END) AS FLOAT) * 100, 
            2
        )
        ELSE 0 
    END AS PERCENTUAL_DEVOLUCOES_NF_PROPRIA,

    -- Cálculo adicional: Percentual de NF Própria sobre o total geral
    CASE 
        WHEN SUM(QUANTIDADE) > 0 
        THEN ROUND(
            CAST(SUM(CASE WHEN EMISSOR = 'OwnEstablishment' THEN QUANTIDADE ELSE 0 END) AS FLOAT) / 
            CAST(SUM(QUANTIDADE) AS FLOAT) * 100, 
            2
        )
        ELSE 0 
    END AS PERCENTUAL_NF_PROPRIA_TOTAL

FROM (
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
        VALOR,
        CFOP,
        DESC_TIPO_OPERACAO,
        CASE
            WHEN COD_ESTABELECIMENTO IN ('R261', 'R221', 'R222', 'R541', 'R591', 'R281', 'R282', 'R283', 'R611', 'R121', 'R831', 'R351', 'R352', 'R461', 'R521') THEN 'AVANÇAR'
            WHEN COD_ESTABELECIMENTO IN ('R201', 'R311', 'R312', 'R313', 'R191', 'R781', 'R301') THEN 'BASE'
            WHEN COD_ESTABELECIMENTO IN ('R031', 'R041', 'R291', 'R292', 'R641', 'R791', 'R801') THEN 'CRESCER'
            WHEN COD_ESTABELECIMENTO IN ('R651', 'R671', 'R681', 'R021', 'R181', 'R691', 'R131', 'R721', 'R751') THEN 'FORTALEZA'
            WHEN COD_ESTABELECIMENTO IN ('R211', 'R341', 'R451', 'R481', 'R711', 'R231', 'R234', 'R471', 'R472', 'R061', 'R531') THEN 'PLANALTO'
            WHEN COD_ESTABELECIMENTO IN ('R071', 'R074', 'R382', 'R501', 'R502', 'R661', 'R701', 'R491', 'R492', 'R241', 'R243', 'R621', 'R761', 'R371', 'R731', 'R821') THEN 'SUL'
            WHEN COD_ESTABELECIMENTO IN ('R011', 'R511', 'R101', 'R811', 'R051', 'R052', 'R161') THEN 'VISÃO'
            ELSE 'CADASTRAR'
        END AS GRUPO,
        LENGTH([CPF/CNPJ]) AS CARACTERES,
        CASE
            WHEN LENGTH([CPF/CNPJ]) = 11 THEN 'FÍSICA'
            WHEN LENGTH([CPF/CNPJ]) = 14 THEN 'JURÍDICA'
            ELSE 'VERIFICAR'
        END AS PESSOA,
        -- Classificação como VENDA ou DEVOLUCAO baseado no CFOP
        CASE
            WHEN CFOP IN ('1.201', '1.202', '1.203', '1.204', '1.410', '1.411', '1.553', '1.660', '1.661', '1.662',
                         '2.201', '2.202', '2.203', '2.204', '2.410', '2.411', '2.553', '2.660', '2.661', '2.662',
                         '3.201', '3.202', '3.211', '3.553') THEN 'DEVOLUCAO'
            WHEN CFOP IN ('5.101', '5.102', '5.103', '5.104', '5.105', '5.106', '5.107', '5.108', '5.109', 
                         '5.110', '5.111', '5.112', '5.113', '5.114', '5.115', '5.116', '5.201', '5.202',
                         '5.203', '5.204', '5.205', '5.206', '5.207', '5.208', '5.209', '5.401', '5.402',
                         '5.403', '5.404', '5.405', '5.501', '5.502', '5.503', '5.504', '6.101', '6.102',
                         '6.103', '6.104', '6.105', '6.106', '6.107', '6.108', '6.109', '6.110', '6.111',
                         '6.112', '6.113', '6.114', '6.115', '6.116') THEN 'VENDA'
            ELSE 'OUTROS'
        END AS TIPO_OPERACAO
    FROM 
        VW_AUDIT_RM_ORDENS_VENDA
    WHERE 
        DATA_NOTA_FISCAL BETWEEN '2025-10-01' AND '2025-10-31' 
        AND PARA_FATURAMENTO = 'Sim'
        AND LENGTH([CPF/CNPJ]) = 14
        AND EMISSOR = 'OwnEstablishment'
        AND COD_ESTABELECIMENTO IN ('R261', 'R221', 'R222', 'R541', 'R591', 'R281', 'R282', 'R283', 'R611', 'R121', 'R831', 'R351', 'R352', 'R461', 'R521')
        AND (CFOP IN ('1.201', '1.202', '1.203', '1.204', '1.410', '1.411', '1.553', '1.660', '1.661', '1.662',
                     '2.201', '2.202', '2.203', '2.204', '2.410', '2.411', '2.553', '2.660', '2.661', '2.662',
                     '3.201', '3.202', '3.211', '3.553')
             OR CFOP IN ('5.101', '5.102', '5.103', '5.104', '5.105', '5.106', '5.107', '5.108', '5.109', 
                        '5.110', '5.111', '5.112', '5.113', '5.114', '5.115', '5.116', '5.201', '5.202',
                        '5.203', '5.204', '5.205', '5.206', '5.207', '5.208', '5.209', '5.401', '5.402',
                        '5.403', '5.404', '5.405', '5.501', '5.502', '5.503', '5.504', '6.101', '6.102',
                        '6.103', '6.104', '6.105', '6.106', '6.107', '6.108', '6.109', '6.110', '6.111',
                        '6.112', '6.113', '6.114', '6.115', '6.116'))
) AS BASE
GROUP BY GRUPO, COD_ESTABELECIMENTO, PESSOA
ORDER BY GRUPO, PERCENTUAL_DEVOLUCOES_NF_PROPRIA DESC;
