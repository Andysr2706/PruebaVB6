/*
   04_permisos.sql
   Usuario de aplicacion con permisos minimos: SOLO EXECUTE sobre los
   procedimientos. No tiene SELECT/INSERT/UPDATE/DELETE sobre tablas:
   accede a los datos unicamente via SP (encadenamiento de propiedad dbo).
   Cambiar la clave antes de ejecutar y ponerla en config.ini
*/
USE master;
GO
IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = N'CobranzaApp')
    CREATE LOGIN CobranzaApp WITH PASSWORD = N'Cobr@nzaApp2026!',
        DEFAULT_DATABASE = Cobranza, CHECK_POLICY = ON, CHECK_EXPIRATION = OFF;
GO
USE Cobranza;
GO
IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'CobranzaApp')
    CREATE USER CobranzaApp FOR LOGIN CobranzaApp;
GO
IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'rol_cobranza_app')
    CREATE ROLE rol_cobranza_app;
GO
ALTER ROLE rol_cobranza_app ADD MEMBER CobranzaApp;
GO
GRANT EXECUTE ON dbo.usp_Cliente_Buscar                TO rol_cobranza_app;
GRANT EXECUTE ON dbo.usp_Factura_PendientesPorCliente  TO rol_cobranza_app;
GRANT EXECUTE ON dbo.usp_Pago_Simular                  TO rol_cobranza_app;
GRANT EXECUTE ON dbo.usp_Pago_Aplicar                  TO rol_cobranza_app;
GRANT EXECUTE ON dbo.usp_Reporte_EstadoCuenta          TO rol_cobranza_app;
GRANT EXECUTE ON dbo.usp_Reporte_EstadoCuenta_Pagos    TO rol_cobranza_app;
GRANT EXECUTE ON dbo.usp_Reporte_Antiguedad            TO rol_cobranza_app;
GRANT EXECUTE ON dbo.usp_Mora_Calcular                 TO rol_cobranza_app;
GRANT EXECUTE ON dbo.usp_Parametro_TasaDefecto         TO rol_cobranza_app;
GRANT EXECUTE ON dbo.usp_LogError_Insertar             TO rol_cobranza_app;
GO
/* Verificacion: debe listar solo EXECUTE */
SELECT pr.name AS Principal, pe.permission_name, OBJECT_NAME(pe.major_id) AS Objeto
FROM sys.database_permissions AS pe
JOIN sys.database_principals AS pr ON pr.principal_id = pe.grantee_principal_id
WHERE pr.name IN (N'rol_cobranza_app', N'CobranzaApp')
ORDER BY Objeto;
GO
