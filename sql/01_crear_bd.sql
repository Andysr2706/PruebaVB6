/*
   01_crear_bd.sql
   Crea la base de datos Cobranza. Ejecutar como administrador (sa o
   cuenta de Windows con sysadmin). Orden: 01 -> 02 -> 03 -> 04 -> 05 -> 06 -> 07
*/
USE master;
GO
IF DB_ID(N'Cobranza') IS NOT NULL
BEGIN
    ALTER DATABASE Cobranza SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE Cobranza;
END
GO
CREATE DATABASE Cobranza COLLATE Modern_Spanish_CI_AS;
GO
/* READ_COMMITTED_SNAPSHOT: las lecturas de reportes no bloquean ni son
   bloqueadas por la aplicacion de pagos. Las escrituras siguen usando
   bloqueos explicitos (ver usp_Pago_Aplicar). */
ALTER DATABASE Cobranza SET READ_COMMITTED_SNAPSHOT ON;
GO
