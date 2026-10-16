/*
   06_datos_volumen.sql
   Genera 7 vendedores, 2,000 clientes y 100,000 facturas (set-based,
   segundos) para medir el rendimiento de los reportes. Re-ejecutable:
   borra antes los datos de volumen (codigos VOL / numeros V).
*/
USE Cobranza;
GO
SET NOCOUNT ON;
DECLARE @Hoy DATE = CAST(GETDATE() AS DATE);

DELETE d FROM dbo.PagoDetalle d JOIN dbo.Factura f ON f.FacturaId = d.FacturaId WHERE f.NumeroFactura LIKE 'V%';
DELETE FROM dbo.Factura WHERE NumeroFactura LIKE 'V%';
DELETE p FROM dbo.Pago p JOIN dbo.Cliente c ON c.ClienteId = p.ClienteId WHERE c.Codigo LIKE 'VOL%';
DELETE FROM dbo.Cliente WHERE Codigo LIKE 'VOL%';

IF NOT EXISTS (SELECT 1 FROM dbo.Vendedor WHERE Nombre LIKE N'Vendedor Vol%')
    INSERT INTO dbo.Vendedor (Nombre)
    SELECT N'Vendedor Vol ' + CAST(n AS NVARCHAR(3))
    FROM (VALUES (1),(2),(3),(4),(5),(6),(7)) AS t(n);

;WITH n AS (
    SELECT TOP (2000) ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS i
    FROM sys.all_objects a CROSS JOIN sys.all_objects b)
INSERT INTO dbo.Cliente (Codigo, Nombre, NIT, VendedorId, LimiteCredito, TasaMoraAnual, DiasGracia)
SELECT  'VOL' + RIGHT('00000' + CAST(i AS VARCHAR(5)), 5),
        N'Cliente Volumen ' + CAST(i AS NVARCHAR(5)),
        CASE WHEN i % 25 = 0 THEN NULL ELSE CAST(1000000 + i AS VARCHAR(10)) + '-' + CAST(i % 10 AS VARCHAR(1)) END,
        CASE WHEN i % 50 = 0 THEN NULL
             ELSE (SELECT MIN(VendedorId) FROM dbo.Vendedor) + (i % (SELECT COUNT(*) FROM dbo.Vendedor)) END,
        CAST(5000 + (i % 20) * 2500 AS DECIMAL(18,2)),
        CASE WHEN i % 10 = 0 THEN NULL ELSE CAST(0.12 + (i % 5) * 0.03 AS DECIMAL(9,6)) END,
        CASE WHEN i % 7 = 0 THEN NULL ELSE i % 6 END
FROM n;

DECLARE @MinCli INT = (SELECT MIN(ClienteId) FROM dbo.Cliente WHERE Codigo LIKE 'VOL%');

;WITH n AS (
    SELECT TOP (100000) ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS i
    FROM sys.all_objects a CROSS JOIN sys.all_objects b CROSS JOIN sys.all_objects c),
g AS (
    SELECT i,
           @MinCli + (ABS(CHECKSUM(i * 7919)) % 2000)                        AS ClienteId,
           DATEADD(DAY, -(ABS(CHECKSUM(i * 104729)) % 400), @Hoy)            AS Emision,
           CAST(100 + (ABS(CHECKSUM(i * 31337)) % 4990000) / 100.0 AS DECIMAL(18,2)) AS Monto
    FROM n)
INSERT INTO dbo.Factura (NumeroFactura, ClienteId, Tipo, FechaEmision, FechaVencimiento, Monto, Saldo)
SELECT  'V' + RIGHT('0000000' + CAST(i AS VARCHAR(7)), 7),
        ClienteId, 'FC', Emision,
        CASE WHEN i % 40 = 0 THEN NULL ELSE DATEADD(DAY, 30, Emision) END,
        Monto,
        CASE WHEN i % 3 = 0 THEN 0                                   -- pagada
             WHEN i % 5 = 0 THEN CAST(ROUND(Monto / 2, 2) AS DECIMAL(18,2))  -- abono parcial
             ELSE Monto END
FROM g;

UPDATE STATISTICS dbo.Factura WITH FULLSCAN;
UPDATE STATISTICS dbo.Cliente WITH FULLSCAN;

SELECT COUNT(*) AS TotalFacturas, SUM(CASE WHEN NumeroFactura LIKE 'V%' THEN 1 ELSE 0 END) AS FacturasVolumen
FROM dbo.Factura;
GO
