/*
   07_evidencia_rendimiento.sql
   Ejecutar en SSMS con "Include Actual Execution Plan" (Ctrl+M).
   Guardar la pestana Messages (tiempos/lecturas) y el plan (.sqlplan)
   en docs/evidencia como evidencia de rendimiento.
*/
USE Cobranza;
GO
DECLARE @Corte DATE = CAST(GETDATE() AS DATE);
SET STATISTICS IO ON;
SET STATISTICS TIME ON;

PRINT '=== 1) usp_Reporte_Antiguedad (con indices) ===';
EXEC dbo.usp_Reporte_Antiguedad @FechaCorte = @Corte;

PRINT '=== 2) Misma consulta FORZANDO scan del indice clustered (sin IX) ===';
SELECT  ISNULL(v.Nombre, N'(SIN VENDEDOR)') AS Vendedor, c.Codigo, SUM(f.Saldo) AS Total
FROM dbo.Factura AS f WITH (INDEX(0))
JOIN dbo.Cliente AS c ON c.ClienteId = f.ClienteId
LEFT JOIN dbo.Vendedor AS v ON v.VendedorId = c.VendedorId
WHERE f.Saldo <> 0 AND f.FechaEmision <= @Corte
GROUP BY v.Nombre, c.Codigo;

PRINT '=== 3) usp_Reporte_EstadoCuenta (seek por cliente) ===';
DECLARE @Cli INT = (SELECT TOP 1 ClienteId FROM dbo.Cliente WHERE Codigo LIKE 'VOL%' ORDER BY ClienteId);
EXEC dbo.usp_Reporte_EstadoCuenta @ClienteId = @Cli, @FechaCorte = @Corte;

SET STATISTICS IO OFF;
SET STATISTICS TIME OFF;
GO
