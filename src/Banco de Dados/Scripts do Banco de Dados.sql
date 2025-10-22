-- =====================================================================
-- Criação do Banco de Dados
-- =====================================================================
    CREATE DATABASE dbFinanceiro;
    GO

    USE dbFinanceiro;
    GO

-- =====================================================================
-- Criação da Tabela: CadastroTransacoes
-- =====================================================================
    CREATE TABLE dbo.CadastroTransacoes
    (
        Id_Transacao       INT IDENTITY(1,1) PRIMARY KEY,               -- Gerado automaticamente
        Numero_Cartao      CHAR(16) NOT NULL,                           -- 16 dígitos fixos
        Valor_Transacao    DECIMAL(18,2) NOT NULL CHECK (Valor_Transacao > 0),  -- Valor positivo
        Data_Transacao     DATETIME NOT NULL DEFAULT GETDATE(),         -- Data/hora do registro
        Descricao          VARCHAR(255) NOT NULL,                           -- Descrição opcional
        Status_Transacao   VARCHAR(20) NOT NULL CHECK (Status_Transacao IN ('Aprovada', 'Pendente', 'Cancelada'))  -- Status permitido
    );
    GO

-- =====================================================================
-- Índice adicional para melhorar consultas por número de cartão
-- =====================================================================
    CREATE NONCLUSTERED INDEX IX_CadastroTransacoes_NumeroCartao
    ON dbo.CadastroTransacoes (Numero_Cartao);
    GO

-- =====================================================================
-- Descrições (metadados)
-- =====================================================================
    EXEC sp_addextendedproperty 
        @name = N'MS_Description', 
        @value = N'Tabela que armazena o cadastro de transações de cartão.', 
        @level0type = N'SCHEMA', @level0name = N'dbo', 
        @level1type = N'TABLE',  @level1name = N'CadastroTransacoes';
    GO

    EXEC sp_addextendedproperty 
        @name = N'MS_Description', 
        @value = N'Status possíveis: Aprovada, Pendente ou Cancelada.', 
        @level0type = N'SCHEMA', @level0name = N'dbo', 
        @level1type = N'TABLE',  @level1name = N'CadastroTransacoes', 
        @level2type = N'COLUMN', @level2name = N'Status_Transacao';
    GO



-- ********** ////////////  ***   \\\\\\\\\\\\\\\\\ **********

-- =====================================================================
-- Stored Procedure — Total de Transações por Período
-- =====================================================================
CREATE PROCEDURE dbo.sp_CalculaTotaisTransacoes
    @Data_Inicial VARCHAR(20),
    @Data_Final VARCHAR(20),
    @Status_Transacao VARCHAR(20) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @DtIni DATETIME = TRY_CONVERT(DATETIME, @Data_Inicial, 103);
    DECLARE @DtFim DATETIME = TRY_CONVERT(DATETIME, @Data_Final, 103);

    IF @DtIni IS NULL OR @DtFim IS NULL
    BEGIN
        RAISERROR('Data inválida. Use o formato DD/MM/YYYY.', 16, 1);
        RETURN;
    END;

    SELECT 
        Numero_Cartao,
        SUM(Valor_Transacao) AS Valor_Total,
        COUNT(*) AS Quantidade_Transacoes,
        Status_Transacao
    FROM dbo.CadastroTransacoes
    WHERE Data_Transacao BETWEEN @DtIni AND @DtFim
      AND (@Status_Transacao IS NULL OR Status_Transacao = @Status_Transacao)
    GROUP BY Numero_Cartao, Status_Transacao
    ORDER BY Valor_Total DESC;
END;
GO

    /*
    --Exemplo de Uso:
    EXEC dbo.sp_CalculaTotaisTransacoes 
        @Data_Inicial = '01/01/2025', 
        @Data_Final = '31/12/2025', 
        @Status_Transacao = 'Aprovada';

    */


-- =====================================================================
-- 2. Função Escalar — Categorização de Valor
-- =====================================================================
    CREATE FUNCTION dbo.fn_CategoriaValor(@Valor DECIMAL(18,2))
    RETURNS VARCHAR(20)
    AS
    BEGIN
        DECLARE @Categoria VARCHAR(20);

        IF @Valor > 2000
            SET @Categoria = 'Premium';
        ELSE IF @Valor BETWEEN 1000 AND 2000
            SET @Categoria = 'Alta';
        ELSE IF @Valor BETWEEN 500 AND 999.99
            SET @Categoria = 'Média';
        ELSE
            SET @Categoria = 'Baixa';

        RETURN @Categoria;
    END;
    GO
    /*
        --Exemplo de Uso:
        SELECT dbo.fn_CategoriaValor(1500) AS Categoria;  -- Retorna 'Alta'
    */


-- =====================================================================
-- 3. Table-Valued Function — Transações Categorizadas por Período
-- =====================================================================
    CREATE FUNCTION dbo.fn_TransacoesCategorizadas
    (
        @Data_Inicial DATETIME,
        @Data_Final DATETIME
    )
    RETURNS TABLE
    AS
    RETURN
    (
        -- Recomenda-se usar a variante abaixo para incluir o dia final inteiro:
        SELECT
            t.Id_Transacao,
            t.Numero_Cartao,
            t.Valor_Transacao,
            t.Data_Transacao,
            t.Status_Transacao,
            dbo.fn_CategoriaValor(t.Valor_Transacao) AS Categoria
        FROM dbo.CadastroTransacoes AS t
        WHERE t.Data_Transacao >= @Data_Inicial
          AND t.Data_Transacao <  DATEADD(DAY, 1, CONVERT(date, @Data_Final))
    );
    GO
    /*
        --Exemplo de Uso:
        SELECT * 
        FROM dbo.fn_TransacoesCategorizadas('20250101', '20251231');
    */

-- =====================================================================
-- 4. View — Consolidação Financeira
-- =====================================================================
    CREATE VIEW dbo.vw_ConsolidadoFinanceiro
    AS
    SELECT
        t.Numero_Cartao,
        COUNT(*) AS Quantidade_Transacoes,
        SUM(t.Valor_Transacao) AS Valor_Total,
        AVG(t.Valor_Transacao) AS Valor_Medio,
        MAX(t.Valor_Transacao) AS Maior_Valor,
        MIN(t.Valor_Transacao) AS Menor_Valor,
        t.Status_Transacao,
        dbo.fn_CategoriaValor(AVG(t.Valor_Transacao)) AS Categoria_Media
    FROM dbo.CadastroTransacoes t
    GROUP BY t.Numero_Cartao, t.Status_Transacao;
    GO
    /*
        --Exemplo de Uso:
        SELECT * FROM dbo.vw_ConsolidadoFinanceiro;
    */




/*
--===================================================================
SCRIPT PARA CRIAR TRANSAÇÕES COM:
Números de cartão aleatórios (16 dígitos)
Valores variados (50 até 5000)
Datas aleatórias nos últimos 6 meses
Status aleatórios: Aprovada, Pendente, Cancelada
--===================================================================
*/

USE dbFinanceiro;
GO

SET NOCOUNT ON;

DECLARE @i INT = 1;

WHILE @i <= 200
BEGIN
    INSERT INTO dbo.CadastroTransacoes
        (Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status_Transacao)
    VALUES
        (
            -- Número de cartão com 16 dígitos (fictício)
            RIGHT('0000000000000000' + CAST(ABS(CHECKSUM(NEWID())) AS VARCHAR(16)), 16),

            -- Valor aleatório entre 50 e 5000 com 2 casas decimais
            CAST(ROUND(RAND(CHECKSUM(NEWID())) * (5000.00 - 50.00) + 50.00, 2) AS DECIMAL(10,2)),

            -- Data aleatória nos últimos 180 dias
            DATEADD(DAY, -ABS(CHECKSUM(NEWID())) % 180, GETDATE()),

            -- Descrição aleatória (variação por CASE) + sufixo com o número do registro
            CASE ABS(CHECKSUM(NEWID())) % 6
                WHEN 0 THEN 'Pagamento fatura - teste ' + CAST(@i AS VARCHAR(6))
                WHEN 1 THEN 'Compra online - teste ' + CAST(@i AS VARCHAR(6))
                WHEN 2 THEN 'Saque ATM - teste ' + CAST(@i AS VARCHAR(6))
                WHEN 3 THEN 'Transferencia - teste ' + CAST(@i AS VARCHAR(6))
                WHEN 4 THEN 'Recarga celular - teste ' + CAST(@i AS VARCHAR(6))
                ELSE      'Serviço X - teste ' + CAST(@i AS VARCHAR(6))
            END,

            -- Status aleatório
            CASE ABS(CHECKSUM(NEWID())) % 3
                WHEN 0 THEN 'Aprovada'
                WHEN 1 THEN 'Pendente'
                ELSE 'Cancelada'
            END
        );

    SET @i += 1;
END;

PRINT 'Inseridos 200 registros (com descricao) na tabela dbo.CadastroTransacoes.';
GO


--===================================================================
--Script: Inserir 30 novas transações por cartão existente
--===================================================================
USE dbFinanceiro;
GO

SET NOCOUNT ON;

DECLARE 
    @NumeroCartao VARCHAR(16),
    @i INT,
    @Valor DECIMAL(10,2),
    @Data DATETIME,
    @Descricao VARCHAR(255),
    @Status VARCHAR(20);

-- Cursor para percorrer todos os cartões existentes
DECLARE Cartoes CURSOR FOR
    SELECT DISTINCT TOP 20 Numero_Cartao FROM dbo.CadastroTransacoes; --LIMITADO A 20 REGISTROS
     
OPEN Cartoes;
FETCH NEXT FROM Cartoes INTO @NumeroCartao;

WHILE @@FETCH_STATUS = 0
BEGIN
    SET @i = 1;
    WHILE @i <= 30
    BEGIN
        -- Gera valores aleatórios
        SET @Valor = ROUND(RAND() * 3000 + 10, 2);  -- de 10 a 3010
        SET @Data = DATEADD(DAY, -ABS(CHECKSUM(NEWID()) % 90), GETDATE()); -- últimos 90 dias
        SET @Descricao = 'Transação automática ' + CAST(@i AS VARCHAR(2)) + ' do cartão ' + @NumeroCartao;

        -- Define status aleatoriamente
        DECLARE @rnd INT = ABS(CHECKSUM(NEWID()) % 3);
        SET @Status = CASE @rnd
                        WHEN 0 THEN 'Aprovada'
                        WHEN 1 THEN 'Pendente'
                        ELSE 'Cancelada'
                      END;

        INSERT INTO dbo.CadastroTransacoes (Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status_Transacao)
        VALUES (@NumeroCartao, @Valor, @Data, @Descricao, @Status);

        SET @i = @i + 1;
    END

    FETCH NEXT FROM Cartoes INTO @NumeroCartao;
END

CLOSE Cartoes;
DEALLOCATE Cartoes;

PRINT 'Inserção concluída: 30 transações criadas para cada cartão existente.';
