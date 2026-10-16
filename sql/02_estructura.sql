/*
   02_estructura.sql
   Tablas, llaves, restricciones e indices.
   Dinero: DECIMAL(18,2). Tasas: DECIMAL(9,6) (0.240000 = 24% anual).
*/
USE Cobranza;
GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

CREATE TABLE dbo.Parametro (
    Clave         VARCHAR(50)    NOT NULL CONSTRAINT PK_Parametro PRIMARY KEY,
    ValorDecimal  DECIMAL(18,6)  NULL,
    ValorTexto    NVARCHAR(200)  NULL,
    Descripcion   NVARCHAR(200)  NOT NULL
);
GO

CREATE TABLE dbo.Vendedor (
    VendedorId    INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_Vendedor PRIMARY KEY,
    Nombre        NVARCHAR(100)  NOT NULL
);
GO

CREATE TABLE dbo.Cliente (
    ClienteId      INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_Cliente PRIMARY KEY,
    Codigo         VARCHAR(20)    NOT NULL CONSTRAINT UQ_Cliente_Codigo UNIQUE,
    Nombre         NVARCHAR(150)  NOT NULL,
    NIT            VARCHAR(20)    NULL,          -- puede faltar (dato incompleto)
    VendedorId     INT            NULL CONSTRAINT FK_Cliente_Vendedor REFERENCES dbo.Vendedor(VendedorId),
    LimiteCredito  DECIMAL(18,2)  NOT NULL CONSTRAINT DF_Cliente_Limite DEFAULT (0),
    TasaMoraAnual  DECIMAL(9,6)   NULL,          -- NULL => usa Parametro TASA_MORA_DEFECTO
    DiasGracia     INT            NULL,          -- NULL => 0
    SaldoAFavor    DECIMAL(18,2)  NOT NULL CONSTRAINT DF_Cliente_SaldoAFavor DEFAULT (0),
    CONSTRAINT CK_Cliente_Limite     CHECK (LimiteCredito >= 0),
    CONSTRAINT CK_Cliente_Tasa       CHECK (TasaMoraAnual IS NULL OR TasaMoraAnual BETWEEN 0 AND 1),
    CONSTRAINT CK_Cliente_Gracia     CHECK (DiasGracia IS NULL OR DiasGracia >= 0),
    CONSTRAINT CK_Cliente_SaldoAFavor CHECK (SaldoAFavor >= 0)
);
GO

/* Tipo FC = factura (monto y saldo positivos); NC = nota de credito
   (monto y saldo negativos). El CHECK impide saldos fuera de rango:
   ningun pago puede dejar una factura con saldo negativo. */
CREATE TABLE dbo.Factura (
    FacturaId        INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_Factura PRIMARY KEY,
    NumeroFactura    VARCHAR(20)    NOT NULL CONSTRAINT UQ_Factura_Numero UNIQUE,
    ClienteId        INT            NOT NULL CONSTRAINT FK_Factura_Cliente REFERENCES dbo.Cliente(ClienteId),
    Tipo             CHAR(2)        NOT NULL CONSTRAINT DF_Factura_Tipo DEFAULT ('FC'),
    FechaEmision     DATE           NOT NULL,
    FechaVencimiento DATE           NULL,        -- puede faltar (dato incompleto)
    Monto            DECIMAL(18,2)  NOT NULL,
    Saldo            DECIMAL(18,2)  NOT NULL,
    CONSTRAINT CK_Factura_Tipo  CHECK (Tipo IN ('FC','NC')),
    CONSTRAINT CK_Factura_Saldo CHECK (
           (Tipo = 'FC' AND Monto > 0 AND Saldo BETWEEN 0 AND Monto)
        OR (Tipo = 'NC' AND Monto < 0 AND Saldo BETWEEN Monto AND 0)),
    CONSTRAINT CK_Factura_Fechas CHECK (FechaVencimiento IS NULL OR FechaVencimiento >= FechaEmision)
);
GO

/* UQ_Pago_Referencia evita registrar dos veces el mismo pago
   (doble clic, reintento o dos cajeros con el mismo comprobante). */
CREATE TABLE dbo.Pago (
    PagoId            INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_Pago PRIMARY KEY,
    ClienteId         INT            NOT NULL CONSTRAINT FK_Pago_Cliente REFERENCES dbo.Cliente(ClienteId),
    Fecha             DATE           NOT NULL,
    Monto             DECIMAL(18,2)  NOT NULL,
    FormaPago         VARCHAR(15)    NOT NULL,
    Referencia        VARCHAR(50)    NOT NULL,
    MontoAplicado     DECIMAL(18,2)  NOT NULL,
    SaldoAFavor       DECIMAL(18,2)  NOT NULL,   -- excedente generado por este pago
    Usuario           NVARCHAR(128)  NOT NULL,
    FechaRegistro     DATETIME2(0)   NOT NULL CONSTRAINT DF_Pago_FechaReg DEFAULT (SYSDATETIME()),
    CONSTRAINT CK_Pago_Monto     CHECK (Monto > 0),
    CONSTRAINT CK_Pago_Forma     CHECK (FormaPago IN ('EFECTIVO','CHEQUE','TRANSFERENCIA','TARJETA','DEPOSITO')),
    CONSTRAINT CK_Pago_Cuadre    CHECK (MontoAplicado >= 0 AND SaldoAFavor >= 0 AND MontoAplicado + SaldoAFavor = Monto),
    CONSTRAINT UQ_Pago_Referencia UNIQUE (ClienteId, FormaPago, Referencia)
);
GO

CREATE TABLE dbo.PagoDetalle (
    PagoDetalleId  INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_PagoDetalle PRIMARY KEY,
    PagoId         INT            NOT NULL CONSTRAINT FK_PagoDetalle_Pago REFERENCES dbo.Pago(PagoId),
    FacturaId      INT            NOT NULL CONSTRAINT FK_PagoDetalle_Factura REFERENCES dbo.Factura(FacturaId),
    MontoAplicado  DECIMAL(18,2)  NOT NULL,
    SaldoAnterior  DECIMAL(18,2)  NOT NULL,
    SaldoNuevo     DECIMAL(18,2)  NOT NULL,
    CONSTRAINT CK_PagoDetalle_Monto  CHECK (MontoAplicado > 0),
    CONSTRAINT CK_PagoDetalle_Cuadre CHECK (SaldoNuevo = SaldoAnterior - MontoAplicado AND SaldoNuevo >= 0),
    CONSTRAINT UQ_PagoDetalle UNIQUE (PagoId, FacturaId)
);
GO

CREATE TABLE dbo.LogError (
    LogErrorId     INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_LogError PRIMARY KEY,
    Fecha          DATETIME2(0)   NOT NULL CONSTRAINT DF_LogError_Fecha DEFAULT (SYSDATETIME()),
    Usuario        NVARCHAR(128)  NOT NULL,
    Equipo         NVARCHAR(128)  NULL,
    Origen         VARCHAR(10)    NOT NULL,     -- VB6 / SQL
    Modulo         NVARCHAR(100)  NOT NULL,
    Procedimiento  NVARCHAR(100)  NOT NULL,
    Numero         INT            NOT NULL,
    Descripcion    NVARCHAR(2000) NOT NULL
);
GO

/* ---------------- Indices (justificacion en DECISIONES.md) ----------
   IX_Factura_Cliente_Venc: soporta pendientes por cliente (seek por
   ClienteId, ya ordenado por vencimiento/numero = orden FIFO) y el
   estado de cuenta, cubriendo las columnas que leen los SP.
   IX_Factura_Saldo: soporta antiguedad de saldos (solo saldo <> 0).   */
CREATE INDEX IX_Factura_Cliente_Venc ON dbo.Factura (ClienteId, FechaVencimiento, NumeroFactura)
    INCLUDE (Tipo, FechaEmision, Monto, Saldo);
CREATE INDEX IX_Factura_Saldo ON dbo.Factura (Saldo)
    INCLUDE (ClienteId, FechaVencimiento, Tipo);
CREATE INDEX IX_Cliente_Vendedor ON dbo.Cliente (VendedorId) INCLUDE (Codigo, Nombre);
CREATE INDEX IX_Cliente_Nombre ON dbo.Cliente (Nombre);
CREATE INDEX IX_Pago_Cliente_Fecha ON dbo.Pago (ClienteId, Fecha);
CREATE INDEX IX_PagoDetalle_Factura ON dbo.PagoDetalle (FacturaId);
GO
