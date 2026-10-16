/*
   05_datos_iniciales.sql
   Fechas RELATIVAS a hoy para que los rangos de antiguedad se cumplan
   sin importar el dia en que se ejecute el script.
   Escenarios:
     C001 sin saldo            C002 saldo al dia
     C003 vencido 1-30         C004 vencido 31-60 (tasa NULL -> defecto)
     C005 vencido 61-90 y >90  C006 con nota de credito (saldo negativo)
     C007 datos incompletos    C008 varias facturas para demostrar FIFO
*/
USE Cobranza;
GO
SET NOCOUNT ON;
DECLARE @Hoy DATE = CAST(GETDATE() AS DATE);

INSERT INTO dbo.Parametro (Clave, ValorDecimal, ValorTexto, Descripcion) VALUES
 ('TASA_MORA_DEFECTO', 0.180000, NULL, N'Tasa de mora anual por defecto (18%)');

INSERT INTO dbo.Vendedor (Nombre) VALUES (N'Ana Morales'), (N'Carlos Pérez'), (N'Lucía Gómez');

INSERT INTO dbo.Cliente (Codigo, Nombre, NIT, VendedorId, LimiteCredito, TasaMoraAnual, DiasGracia) VALUES
 ('C001', N'Distribuidora El Sol',        '1234567-8', 1, 10000.00, 0.240000, 5),
 ('C002', N'Ferretería La Económica',     '2345678-9', 1, 25000.00, 0.240000, 0),
 ('C003', N'Comercial Peñate',            '3456789-1', 2, 15000.00, 0.240000, 0),
 ('C004', N'Agroinsumos del Norte',       '4567891-2', 2, 30000.00, NULL,     3),
 ('C005', N'Importadora Ñandú',           '5678912-3', 3, 50000.00, 0.300000, NULL),
 ('C006', N'Librería San José',           '6789123-4', 3,  8000.00, 0.240000, 0),
 ('C007', N'Cliente Datos Incompletos',   NULL,        NULL, 5000.00, NULL,   NULL),
 ('C008', N'Supermercados La Bendición',  '7891234-5', 1, 40000.00, 0.200000, 2);

DECLARE @C001 INT = (SELECT ClienteId FROM dbo.Cliente WHERE Codigo='C001'),
        @C002 INT = (SELECT ClienteId FROM dbo.Cliente WHERE Codigo='C002'),
        @C003 INT = (SELECT ClienteId FROM dbo.Cliente WHERE Codigo='C003'),
        @C004 INT = (SELECT ClienteId FROM dbo.Cliente WHERE Codigo='C004'),
        @C005 INT = (SELECT ClienteId FROM dbo.Cliente WHERE Codigo='C005'),
        @C006 INT = (SELECT ClienteId FROM dbo.Cliente WHERE Codigo='C006'),
        @C007 INT = (SELECT ClienteId FROM dbo.Cliente WHERE Codigo='C007'),
        @C008 INT = (SELECT ClienteId FROM dbo.Cliente WHERE Codigo='C008');

INSERT INTO dbo.Factura (NumeroFactura, ClienteId, Tipo, FechaEmision, FechaVencimiento, Monto, Saldo) VALUES
 -- C001 sin saldo (factura totalmente pagada)
 ('F-0001', @C001, 'FC', DATEADD(DAY,-60,@Hoy), DATEADD(DAY,-30,@Hoy),  1500.00,    0.00),
 -- C002 al dia
 ('F-0002', @C002, 'FC', DATEADD(DAY,-10,@Hoy), DATEADD(DAY, 20,@Hoy),  3200.50, 3200.50),
 ('F-0003', @C002, 'FC', DATEADD(DAY, -5,@Hoy), DATEADD(DAY, 25,@Hoy),  1800.00, 1800.00),
 -- C003 vencido 1-30
 ('F-0004', @C003, 'FC', DATEADD(DAY,-45,@Hoy), DATEADD(DAY,-15,@Hoy),  5000.00, 5000.00),
 ('F-0005', @C003, 'FC', DATEADD(DAY,-50,@Hoy), DATEADD(DAY,-20,@Hoy),  2000.00, 2000.00),
 -- C004 vencido 31-60, tasa NULL (usa defecto)
 ('F-0006', @C004, 'FC', DATEADD(DAY,-75,@Hoy), DATEADD(DAY,-45,@Hoy), 12500.75, 12500.75),
 -- C005 vencido 61-90 y mas de 90, dias de gracia NULL
 ('F-0007', @C005, 'FC', DATEADD(DAY,-105,@Hoy), DATEADD(DAY,-75,@Hoy), 8000.00, 8000.00),
 ('F-0008', @C005, 'FC', DATEADD(DAY,-150,@Hoy), DATEADD(DAY,-120,@Hoy),20000.00,15000.00),
 -- C006 con nota de credito (saldo negativo, no genera interes)
 ('F-0009', @C006, 'FC', DATEADD(DAY,-50,@Hoy), DATEADD(DAY,-20,@Hoy),  2500.00, 2500.00),
 ('NC-0001',@C006, 'NC', DATEADD(DAY,-10,@Hoy), NULL,                   -500.00, -500.00),
 -- C007 datos incompletos: sin vencimiento + vencida, sin tasa ni gracia
 ('F-0010', @C007, 'FC', DATEADD(DAY,-40,@Hoy), NULL,                   1200.00, 1200.00),
 ('F-0011', @C007, 'FC', DATEADD(DAY,-70,@Hoy), DATEADD(DAY,-40,@Hoy),   900.00,  900.00),
 -- C008 FIFO: dos facturas con el MISMO vencimiento (desempata el numero)
 ('F-0014', @C008, 'FC', DATEADD(DAY,-60,@Hoy), DATEADD(DAY,-30,@Hoy),  1000.00, 1000.00),
 ('F-0012', @C008, 'FC', DATEADD(DAY,-60,@Hoy), DATEADD(DAY,-30,@Hoy),  1500.00, 1500.00),
 ('F-0013', @C008, 'FC', DATEADD(DAY,-90,@Hoy), DATEADD(DAY,-60,@Hoy),  2000.00, 2000.00),
 ('F-0015', @C008, 'FC', DATEADD(DAY,-20,@Hoy), DATEADD(DAY, 10,@Hoy),  3000.00, 3000.00);

/* Pagos historicos (via el mismo SP) para el subreporte de 90 dias */
DECLARE @PagoId INT, @SaF DECIMAL(18,2), @Rc INT, @F DATE;
SET @F = DATEADD(DAY,-12,@Hoy);
EXEC @Rc = dbo.usp_Pago_Aplicar @ClienteId=@C003, @Fecha=@F, @Monto=1500.00,
     @FormaPago='TRANSFERENCIA', @Referencia='TRF-1001', @Usuario=N'carga_inicial',
     @PagoId=@PagoId OUTPUT, @SaldoAFavor=@SaF OUTPUT;
SET @F = DATEADD(DAY,-35,@Hoy);
EXEC @Rc = dbo.usp_Pago_Aplicar @ClienteId=@C005, @Fecha=@F, @Monto=3000.00,
     @FormaPago='CHEQUE', @Referencia='CHQ-5501', @Usuario=N'carga_inicial',
     @PagoId=@PagoId OUTPUT, @SaldoAFavor=@SaF OUTPUT;

SELECT Codigo, Nombre, SaldoAFavor FROM dbo.Cliente ORDER BY Codigo;
EXEC dbo.usp_Cliente_Buscar;
GO
