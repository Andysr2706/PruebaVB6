/*
   03_procedimientos.sql
   Funciones y procedimientos almacenados.
   REGLA DE MORA: vive en dbo.tvf_Mora (unica fuente en SQL). Los SP y
   ambos reportes la usan; clsCalculoMora (VB6) replica exactamente la
   misma aritmetica y frmPruebasMora compara ambos resultados.
*/
USE Cobranza;
GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

/*
   tvf_Mora: dias de atraso, dias de mora, interes y senal de una factura
   interes = saldo * (tasa / 365) * diasMora, redondeo comercial a 2 dec.
   - Se multiplica primero (exacto en DECIMAL) y se divide al final con
     19 decimales: evita perdida de precision antes del redondeo.
   - ROUND sobre DECIMAL redondea 0.5 alejandose de cero (comercial).
*/
CREATE OR ALTER FUNCTION dbo.tvf_Mora
(
    @Saldo      DECIMAL(18,2),
    @Tasa       DECIMAL(9,6),   -- tasa efectiva (cliente o defecto); NULL = sin tasa
    @DiasGracia INT,            -- NULL = 0
    @FechaVenc  DATE,           -- NULL = sin vencimiento
    @FechaCorte DATE
)
RETURNS TABLE
AS
RETURN
    SELECT  d.DiasAtraso,
            d.DiasMora,
            CAST(CASE WHEN @Saldo IS NULL OR @Saldo <= 0 OR @FechaVenc IS NULL
                           OR @Tasa IS NULL OR d.DiasMora = 0 THEN 0
                      ELSE ROUND(CAST(@Saldo * @Tasa * d.DiasMora AS DECIMAL(26,8))
                                 / CAST(365 AS DECIMAL(10,0)), 2)
                 END AS DECIMAL(18,2)) AS Interes,
            CAST(CASE WHEN @Saldo IS NULL      THEN ''
                      WHEN @Saldo < 0          THEN 'NOTA CREDITO'
                      WHEN @FechaVenc IS NULL  THEN 'SIN VENCIMIENTO'
                      WHEN @Tasa IS NULL       THEN 'SIN TASA'
                      ELSE '' END AS VARCHAR(20)) AS Senal
    FROM (SELECT
            CASE WHEN @FechaVenc IS NULL OR @Saldo IS NULL OR @Saldo <= 0 THEN 0
                 WHEN DATEDIFF(DAY, @FechaVenc, @FechaCorte) > 0
                      THEN DATEDIFF(DAY, @FechaVenc, @FechaCorte) ELSE 0 END AS DiasAtraso,
            CASE WHEN @FechaVenc IS NULL OR @Saldo IS NULL OR @Saldo <= 0 THEN 0
                 WHEN DATEDIFF(DAY, @FechaVenc, @FechaCorte) - ISNULL(@DiasGracia, 0) > 0
                      THEN DATEDIFF(DAY, @FechaVenc, @FechaCorte) - ISNULL(@DiasGracia, 0)
                 ELSE 0 END AS DiasMora) AS d;
GO

/*
   tvf_PagoDistribucion: reparto FIFO de un monto sobre las facturas
   pendientes (vencimiento asc, empate por numero; sin vencimiento al
   final). Unica fuente para la vista previa y para la aplicacion real.
*/
CREATE OR ALTER FUNCTION dbo.tvf_PagoDistribucion
(
    @ClienteId INT,
    @Monto     DECIMAL(18,2)
)
RETURNS TABLE
AS
RETURN
    WITH p AS (
        SELECT  f.FacturaId, f.NumeroFactura, f.FechaVencimiento, f.Saldo,
                ROW_NUMBER() OVER (ORDER BY CASE WHEN f.FechaVencimiento IS NULL THEN 1 ELSE 0 END,
                                            f.FechaVencimiento, f.NumeroFactura) AS Orden,
                SUM(f.Saldo) OVER (ORDER BY CASE WHEN f.FechaVencimiento IS NULL THEN 1 ELSE 0 END,
                                            f.FechaVencimiento, f.NumeroFactura
                                   ROWS UNBOUNDED PRECEDING) AS Acumulado
        FROM dbo.Factura AS f
        WHERE f.ClienteId = @ClienteId AND f.Tipo = 'FC' AND f.Saldo > 0
    )
    SELECT  Orden, FacturaId, NumeroFactura, FechaVencimiento,
            Saldo AS SaldoAnterior,
            CAST(CASE WHEN Acumulado <= @Monto              THEN Saldo
                      WHEN Acumulado - Saldo < @Monto        THEN @Monto - (Acumulado - Saldo)
                      ELSE 0 END AS DECIMAL(18,2)) AS MontoAplicado
    FROM p;
GO

/* NOTA: los parametros y columnas de fecha que cruzan hacia VB6/Crystal
   usan DATETIME porque el proveedor SQLOLEDB (incluido en Windows) no
   reconoce el tipo DATE y lo entregaria como texto. Las tablas usan DATE. */
-- Soporte
CREATE OR ALTER PROCEDURE dbo.usp_LogError_Insertar
    @Usuario       NVARCHAR(128),
    @Equipo        NVARCHAR(128),
    @Modulo        NVARCHAR(100),
    @Procedimiento NVARCHAR(100),
    @Numero        INT,
    @Descripcion   NVARCHAR(2000)
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO dbo.LogError (Usuario, Equipo, Origen, Modulo, Procedimiento, Numero, Descripcion)
    VALUES (ISNULL(@Usuario, SUSER_SNAME()), @Equipo, 'VB6', ISNULL(@Modulo, ''),
            ISNULL(@Procedimiento, ''), ISNULL(@Numero, 0), ISNULL(@Descripcion, ''));
    RETURN 0;
END
GO

CREATE OR ALTER PROCEDURE dbo.usp_Parametro_TasaDefecto
    @Tasa DECIMAL(9,6) OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    SELECT @Tasa = CAST(ValorDecimal AS DECIMAL(9,6))
    FROM dbo.Parametro WHERE Clave = 'TASA_MORA_DEFECTO';
    RETURN 0;
END
GO

/* Usado por frmPruebasMora para comparar VB6 vs SQL con los mismos datos */
CREATE OR ALTER PROCEDURE dbo.usp_Mora_Calcular
    @Saldo      DECIMAL(18,2),
    @Tasa       DECIMAL(9,6),
    @DiasGracia INT,
    @FechaVenc  DATETIME,
    @FechaCorte DATETIME,
    @DiasMora   INT            OUTPUT,
    @Interes    DECIMAL(18,2)  OUTPUT,
    @Senal      VARCHAR(20)    OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    SELECT @DiasMora = DiasMora, @Interes = Interes, @Senal = Senal
    FROM dbo.tvf_Mora(@Saldo, @Tasa, @DiasGracia, @FechaVenc, @FechaCorte);
    RETURN 0;
END
GO

-- Consultas
CREATE OR ALTER PROCEDURE dbo.usp_Cliente_Buscar
    @Texto      NVARCHAR(100) = NULL,   -- busqueda parcial por codigo, nombre o NIT
    @FechaCorte DATETIME          = NULL
AS
BEGIN
    SET NOCOUNT ON;
    IF @FechaCorte IS NULL SET @FechaCorte = GETDATE();
    SET @FechaCorte = CAST(CAST(@FechaCorte AS DATE) AS DATETIME);  -- sin hora
    SET @Texto = NULLIF(LTRIM(RTRIM(@Texto)), N'');

    DECLARE @TasaDef DECIMAL(9,6);
    SELECT @TasaDef = CAST(ValorDecimal AS DECIMAL(9,6)) FROM dbo.Parametro WHERE Clave = 'TASA_MORA_DEFECTO';

    SELECT  c.ClienteId, c.Codigo, c.Nombre, c.NIT, c.LimiteCredito,
            ISNULL(SUM(f.Saldo), 0)                                              AS SaldoTotal,
            ISNULL(SUM(CASE WHEN m.DiasAtraso > 0 AND f.Saldo > 0 THEN f.Saldo ELSE 0 END), 0) AS SaldoVencido,
            ISNULL(MAX(CASE WHEN f.Saldo > 0 THEN m.DiasAtraso END), 0)          AS DiasMayorAtraso,
            ISNULL(SUM(m.Interes), 0)                                            AS InteresMoratorio,
            c.SaldoAFavor,
            CASE WHEN SUM(CASE WHEN f.FacturaId IS NOT NULL AND m.Senal IN ('SIN VENCIMIENTO','SIN TASA')
                               THEN 1 ELSE 0 END) > 0
                 THEN 'REVISAR DATOS' ELSE '' END                                AS Observacion
    FROM dbo.Cliente AS c
    LEFT JOIN dbo.Factura AS f ON f.ClienteId = c.ClienteId AND f.Saldo <> 0
                              AND f.FechaEmision <= @FechaCorte
    OUTER APPLY dbo.tvf_Mora(f.Saldo, COALESCE(c.TasaMoraAnual, @TasaDef), c.DiasGracia,
                             f.FechaVencimiento, @FechaCorte) AS m
    WHERE @Texto IS NULL
       OR c.Codigo LIKE N'%' + @Texto + N'%'
       OR c.Nombre LIKE N'%' + @Texto + N'%'
       OR c.NIT    LIKE N'%' + @Texto + N'%'
    GROUP BY c.ClienteId, c.Codigo, c.Nombre, c.NIT, c.LimiteCredito, c.SaldoAFavor
    ORDER BY c.Nombre;
    RETURN 0;
END
GO

CREATE OR ALTER PROCEDURE dbo.usp_Factura_PendientesPorCliente
    @ClienteId  INT,
    @FechaCorte DATETIME = NULL
AS
BEGIN
    SET NOCOUNT ON;
    IF @FechaCorte IS NULL SET @FechaCorte = GETDATE();
    SET @FechaCorte = CAST(CAST(@FechaCorte AS DATE) AS DATETIME);  -- sin hora

    DECLARE @TasaDef DECIMAL(9,6);
    SELECT @TasaDef = CAST(ValorDecimal AS DECIMAL(9,6)) FROM dbo.Parametro WHERE Clave = 'TASA_MORA_DEFECTO';

    SELECT  f.FacturaId, f.NumeroFactura, f.Tipo,
            CAST(f.FechaEmision AS DATETIME) AS FechaEmision,
            CAST(f.FechaVencimiento AS DATETIME) AS FechaVencimiento,
            f.Monto, f.Saldo,
            COALESCE(c.TasaMoraAnual, @TasaDef) AS TasaAplicada,
            c.DiasGracia,
            m.DiasAtraso, m.DiasMora, m.Interes, m.Senal
    FROM dbo.Factura AS f
    JOIN dbo.Cliente AS c ON c.ClienteId = f.ClienteId
    CROSS APPLY dbo.tvf_Mora(f.Saldo, COALESCE(c.TasaMoraAnual, @TasaDef), c.DiasGracia,
                             f.FechaVencimiento, @FechaCorte) AS m
    WHERE f.ClienteId = @ClienteId AND f.Saldo <> 0
    ORDER BY CASE WHEN f.FechaVencimiento IS NULL THEN 1 ELSE 0 END,
             f.FechaVencimiento, f.NumeroFactura;
    RETURN 0;
END
GO

-- Pagos
CREATE OR ALTER PROCEDURE dbo.usp_Pago_Simular
    @ClienteId   INT,
    @Monto       DECIMAL(18,2),
    @SaldoAFavor DECIMAL(18,2) OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    SELECT @SaldoAFavor = @Monto - ISNULL(SUM(MontoAplicado), 0)
    FROM dbo.tvf_PagoDistribucion(@ClienteId, @Monto);

    SELECT Orden, FacturaId, NumeroFactura, CAST(FechaVencimiento AS DATETIME) AS FechaVencimiento, SaldoAnterior, MontoAplicado,
           CAST(SaldoAnterior - MontoAplicado AS DECIMAL(18,2)) AS SaldoNuevo
    FROM dbo.tvf_PagoDistribucion(@ClienteId, @Monto)
    ORDER BY Orden;
    RETURN 0;
END
GO

/*
   usp_Pago_Aplicar
   Valor de retorno:  0 OK | -1 monto invalido | -2 fecha invalida
                     -3 forma de pago invalida | -4 referencia vacia
                     -5 cliente no existe | -6 pago duplicado
   Errores inesperados: ROLLBACK + LogError + THROW (la app los registra).
   Concurrencia: UPDLOCK+HOLDLOCK sobre la fila del cliente serializa los
   pagos del MISMO cliente (dos cajeros esperan turno); clientes distintos
   no se bloquean entre si. Ademas se verifica que el saldo leido no haya
   cambiado (defensa adicional) y los CHECK impiden saldos negativos.
*/
CREATE OR ALTER PROCEDURE dbo.usp_Pago_Aplicar
    @ClienteId   INT,
    @Fecha       DATETIME,
    @Monto       DECIMAL(18,2),
    @FormaPago   VARCHAR(15),
    @Referencia  VARCHAR(50),
    @Usuario     NVARCHAR(128),
    @PagoId      INT            OUTPUT,
    @SaldoAFavor DECIMAL(18,2)  OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    SET @PagoId = NULL;
    SET @SaldoAFavor = 0;

    IF @Monto IS NULL OR @Monto <= 0 RETURN -1;
    IF @Fecha IS NULL OR CAST(@Fecha AS DATE) > CAST(GETDATE() AS DATE) RETURN -2;
    IF @FormaPago IS NULL OR @FormaPago NOT IN ('EFECTIVO','CHEQUE','TRANSFERENCIA','TARJETA','DEPOSITO') RETURN -3;
    SET @Referencia = LTRIM(RTRIM(@Referencia));
    IF @Referencia IS NULL OR @Referencia = '' RETURN -4;

    DECLARE @Aplic TABLE (
        FacturaId     INT PRIMARY KEY,
        SaldoAnterior DECIMAL(18,2) NOT NULL,
        MontoAplicado DECIMAL(18,2) NOT NULL);
    DECLARE @TotalAplicado DECIMAL(18,2), @Filas INT;

    BEGIN TRY
        BEGIN TRAN;

        -- 1) Bloqueo exclusivo por cliente (serializa cajeros concurrentes)
        IF NOT EXISTS (SELECT 1 FROM dbo.Cliente WITH (UPDLOCK, HOLDLOCK) WHERE ClienteId = @ClienteId)
        BEGIN
            ROLLBACK TRAN;
            RETURN -5;
        END

        -- 2) Idempotencia: el mismo comprobante no se aplica dos veces
        IF EXISTS (SELECT 1 FROM dbo.Pago WHERE ClienteId = @ClienteId
                   AND FormaPago = @FormaPago AND Referencia = @Referencia)
        BEGIN
            ROLLBACK TRAN;
            RETURN -6;
        END

        -- 3) Reparto FIFO
        INSERT INTO @Aplic (FacturaId, SaldoAnterior, MontoAplicado)
        SELECT FacturaId, SaldoAnterior, MontoAplicado
        FROM dbo.tvf_PagoDistribucion(@ClienteId, @Monto)
        WHERE MontoAplicado > 0;

        SELECT @TotalAplicado = ISNULL(SUM(MontoAplicado), 0) FROM @Aplic;
        SET @SaldoAFavor = @Monto - @TotalAplicado;

        -- 4) Encabezado del pago
        INSERT INTO dbo.Pago (ClienteId, Fecha, Monto, FormaPago, Referencia,
                              MontoAplicado, SaldoAFavor, Usuario)
        VALUES (@ClienteId, @Fecha, @Monto, @FormaPago, @Referencia,
                @TotalAplicado, @SaldoAFavor, ISNULL(@Usuario, SUSER_SNAME()));
        SET @PagoId = SCOPE_IDENTITY();

        -- 5) Rebajar saldos solo si no cambiaron desde la lectura
        UPDATE f
           SET f.Saldo = f.Saldo - a.MontoAplicado
        FROM dbo.Factura AS f
        JOIN @Aplic AS a ON a.FacturaId = f.FacturaId
        WHERE f.Saldo = a.SaldoAnterior;
        SET @Filas = @@ROWCOUNT;

        IF @Filas <> (SELECT COUNT(*) FROM @Aplic)
            THROW 50001, N'El saldo de una factura cambio durante la aplicacion del pago.', 1;

        -- 6) Detalle de aplicacion
        INSERT INTO dbo.PagoDetalle (PagoId, FacturaId, MontoAplicado, SaldoAnterior, SaldoNuevo)
        SELECT @PagoId, FacturaId, MontoAplicado, SaldoAnterior, SaldoAnterior - MontoAplicado
        FROM @Aplic;

        -- 7) Excedente como saldo a favor
        IF @SaldoAFavor > 0
            UPDATE dbo.Cliente SET SaldoAFavor = SaldoAFavor + @SaldoAFavor
            WHERE ClienteId = @ClienteId;

        COMMIT TRAN;
        RETURN 0;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRAN;
        SET @PagoId = NULL;
        SET @SaldoAFavor = 0;

        IF ERROR_NUMBER() IN (2601, 2627) RETURN -6;   -- duplicado por carrera

        INSERT INTO dbo.LogError (Usuario, Equipo, Origen, Modulo, Procedimiento, Numero, Descripcion)
        VALUES (ISNULL(@Usuario, SUSER_SNAME()), HOST_NAME(), 'SQL', 'usp_Pago_Aplicar',
                ISNULL(ERROR_PROCEDURE(), 'usp_Pago_Aplicar'), ERROR_NUMBER(), ERROR_MESSAGE());
        THROW;
    END CATCH
END
GO

-- Reportes
/* Estado de cuenta (modelo PULL: Crystal ejecuta este SP) */
CREATE OR ALTER PROCEDURE dbo.usp_Reporte_EstadoCuenta
    @ClienteId  INT,
    @FechaCorte DATETIME
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @TasaDef DECIMAL(9,6);
    SELECT @TasaDef = CAST(ValorDecimal AS DECIMAL(9,6)) FROM dbo.Parametro WHERE Clave = 'TASA_MORA_DEFECTO';

    SELECT  c.ClienteId, c.Codigo, c.Nombre,
            ISNULL(c.NIT, 'SIN NIT')                         AS NIT,
            ISNULL(v.Nombre, N'(SIN VENDEDOR)')              AS Vendedor,
            c.LimiteCredito,
            COALESCE(c.TasaMoraAnual, @TasaDef)              AS TasaAplicada,
            ISNULL(c.DiasGracia, 0)                          AS DiasGracia,
            c.SaldoAFavor,
            CAST(CAST(@FechaCorte AS DATE) AS DATETIME)      AS FechaCorte,
            f.FacturaId, f.NumeroFactura, f.Tipo,
            CAST(f.FechaEmision AS DATETIME)     AS FechaEmision,
            CAST(f.FechaVencimiento AS DATETIME) AS FechaVencimiento,
            f.Monto, f.Saldo,
            ISNULL(m.DiasMora, 0)                            AS DiasMora,
            ISNULL(m.Interes, 0)                             AS Interes,
            ISNULL(m.Senal, '')                              AS Senal
    FROM dbo.Cliente AS c
    LEFT JOIN dbo.Vendedor AS v ON v.VendedorId = c.VendedorId
    LEFT JOIN dbo.Factura  AS f ON f.ClienteId = c.ClienteId AND f.Saldo <> 0
                               AND f.FechaEmision <= @FechaCorte
    OUTER APPLY dbo.tvf_Mora(f.Saldo, COALESCE(c.TasaMoraAnual, @TasaDef), c.DiasGracia,
                             f.FechaVencimiento, @FechaCorte) AS m
    WHERE c.ClienteId = @ClienteId
    ORDER BY CASE WHEN f.FechaVencimiento IS NULL THEN 1 ELSE 0 END,
             f.FechaVencimiento, f.NumeroFactura;
    RETURN 0;
END
GO

/* Subreporte: pagos aplicados en los ultimos 90 dias a la fecha de corte */
CREATE OR ALTER PROCEDURE dbo.usp_Reporte_EstadoCuenta_Pagos
    @ClienteId  INT,
    @FechaCorte DATETIME
AS
BEGIN
    SET NOCOUNT ON;
    SELECT  p.PagoId, CAST(p.Fecha AS DATETIME) AS Fecha, p.FormaPago, p.Referencia, p.Monto AS MontoPago,
            ISNULL(f.NumeroFactura, '(SALDO A FAVOR)') AS NumeroFactura,
            ISNULL(d.MontoAplicado, p.SaldoAFavor)      AS MontoAplicado
    FROM dbo.Pago AS p
    LEFT JOIN dbo.PagoDetalle AS d ON d.PagoId = p.PagoId
    LEFT JOIN dbo.Factura     AS f ON f.FacturaId = d.FacturaId
    WHERE p.ClienteId = @ClienteId
      AND p.Fecha >  DATEADD(DAY, -90, @FechaCorte)
      AND p.Fecha <= @FechaCorte
    ORDER BY p.Fecha, p.PagoId, f.NumeroFactura;
    RETURN 0;
END
GO

/* Antiguedad de saldos (modelo PUSH: VB6 obtiene el Recordset y lo
   entrega al reporte). Agregado por cliente en el servidor para no
   transferir 100k filas a la estacion; Crystal agrupa por vendedor y
   cliente y calcula subtotales. Rangos por dias de atraso (sin gracia). */
CREATE OR ALTER PROCEDURE dbo.usp_Reporte_Antiguedad
    @FechaCorte DATETIME
AS
BEGIN
    SET NOCOUNT ON;
    SELECT  ISNULL(v.VendedorId, 0)              AS VendedorId,
            ISNULL(v.Nombre, N'(SIN VENDEDOR)')  AS Vendedor,
            c.ClienteId, c.Codigo, c.Nombre       AS Cliente,
            SUM(CASE WHEN x.Dias <= 0                  THEN f.Saldo ELSE 0 END) AS AlDia,
            SUM(CASE WHEN x.Dias BETWEEN 1  AND 30     THEN f.Saldo ELSE 0 END) AS D1_30,
            SUM(CASE WHEN x.Dias BETWEEN 31 AND 60     THEN f.Saldo ELSE 0 END) AS D31_60,
            SUM(CASE WHEN x.Dias BETWEEN 61 AND 90     THEN f.Saldo ELSE 0 END) AS D61_90,
            SUM(CASE WHEN x.Dias > 90                  THEN f.Saldo ELSE 0 END) AS Mas90,
            SUM(f.Saldo)                                                        AS Total
    FROM dbo.Factura AS f
    JOIN dbo.Cliente AS c ON c.ClienteId = f.ClienteId
    LEFT JOIN dbo.Vendedor AS v ON v.VendedorId = c.VendedorId
    -- Sin vencimiento o nota de credito => columna "Al dia"
    CROSS APPLY (SELECT CASE WHEN f.FechaVencimiento IS NULL OR f.Saldo < 0 THEN 0
                             ELSE DATEDIFF(DAY, f.FechaVencimiento, @FechaCorte) END AS Dias) AS x
    WHERE f.Saldo <> 0
      AND f.FechaEmision <= @FechaCorte
    GROUP BY v.VendedorId, v.Nombre, c.ClienteId, c.Codigo, c.Nombre
    ORDER BY Vendedor, Cliente;
    RETURN 0;
END
GO
