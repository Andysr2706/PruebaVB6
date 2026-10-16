# Decisiones técnicas: Cobranza Express

## Arquitectura en capas
`frm*` (presentación) → `clsCobranza` / `clsCalculoMora` (negocio) → `clsDataAccess` (datos) → procedimientos almacenados.
- Los formularios no tienen SQL ni `ADODB.Connection`.
- `clsDataAccess` es el único lugar que crea `Connection`, `Command` y `Recordset`.
- `clsCalculoMora` no depende de la base de datos ni de la interfaz, así que se puede probar sola (`frmPruebasMora`).
- Soporte: `modConfig` (config.ini), `modLog` (archivo + `LogError`), `modUtil` (Null, fechas y montos), `modReportes` (Crystal).

## Acceso a datos
- **Solo procedimientos almacenados** con `adCmdStoredProc` y parámetros tipados (`CreateParameter`). Nunca se concatena SQL.
- `RETURN_VALUE` + parámetros `OUTPUT` (`usp_Pago_Aplicar` devuelve `@PagoId` y `@SaldoAFavor`).
- **Cursores según el uso:**
  - Grids: Recordset **desconectado** (`adUseClient`, estático, `ActiveConnection = Nothing`).
  - Lista de clientes para combos: **forward-only / read-only** del servidor, recorrido una sola vez.
- Null manejado de forma explícita (`NzCur`, `NzStr`, `NzLng`, `.Value` al leer campos).
- `ConnectionTimeout` y `CommandTimeout` vienen de config.ini. Si hay un timeout, el usuario ve un mensaje claro.
- La colección `Errors` del proveedor se guarda en el log.
- El usuario `CobranzaApp` solo tiene **EXECUTE** sobre los SP. No tiene acceso a las tablas.
- **Proveedor `SQLOLEDB`** (viene con Windows, no requiere instalar nada). No reconoce el tipo `DATE`, por eso los parámetros y columnas de fecha que salen a VB6 y Crystal son `DATETIME`. Las tablas usan `DATE`.

## Dinero: por qué no Double
`Double` es binario y no representa exactamente 0.10. Las sumas y restas acumulan error y un saldo puede quedar en 0.0000001.
- En SQL se usa `DECIMAL(18,2)`; las tasas, `DECIMAL(9,6)`.
- En VB6 se usa `Currency` (entero escalado, exacto a 4 decimales) y `Decimal` (`CDec`, 28 dígitos) para los cálculos intermedios.
- Los montos capturados se convierten sin `CDbl` ni `Val` (`ParseMonto`), así que no dependen de la región.

## Interés moratorio: dónde vive la regla y cómo se garantiza la consistencia
- **Fuente de verdad: `dbo.tvf_Mora`** (función en línea).
  - La usan `usp_Cliente_Buscar`, `usp_Factura_PendientesPorCliente` y `usp_Reporte_EstadoCuenta`.
  - Crystal **no recalcula**: muestra el `Interes` que entrega el SP. El reporte y SQL dan lo mismo por construcción.
- **VB6 (pantalla): `clsCalculoMora`** replica la misma aritmética:
  1. `saldo × tasa × días` (exacto).
  2. `/ 365` con al menos 19 decimales.
  3. Redondeo comercial a 2 decimales.
  - En SQL, `ROUND` sobre `DECIMAL` ya redondea 0.5 alejándose de cero.
  - En VB6 **no se usa `Round()`**, porque es bancario (2.125 → 2.12). Se usa `Int(x × 100 + 0.5) / 100` sobre el valor absoluto, con `Decimal`.
- **Cómo se comprueba:**
  - `frmClientes` muestra para cada factura *Interés VB6*, *Interés SQL* y la columna *Coincide*.
  - `frmPruebasMora` ejecuta 12 casos (esperado vs. VB6 vs. SQL) con fecha de corte fija: gracia, gracia Null, vencimiento Null, tasa Null (con y sin tasa por defecto), nota de crédito, no vencida, 2.125, 2.135 y 0.005.
- **Reglas de datos incompletos:**
  - Tasa del cliente Null → se usa `Parametro.TASA_MORA_DEFECTO`. Si ambas faltan, no hay interés y se marca "SIN TASA".
  - Gracia Null → 0.
  - Vencimiento Null → no hay interés y se marca "SIN VENCIMIENTO".
  - Saldo negativo (nota de crédito) → no genera interés.

## Transacciones y concurrencia
`usp_Pago_Aplicar` hace todo en una sola transacción (`XACT_ABORT ON`, `TRY/CATCH`; en error: `ROLLBACK` + `LogError` + `THROW`).
- **Bloqueo pesimista por cliente:** `SELECT ... FROM Cliente WITH (UPDLOCK, HOLDLOCK)`. Dos cajeros con el mismo cliente se ponen en fila: el segundo espera y luego ve los saldos ya actualizados. Clientes distintos no se bloquean entre sí.
  - Se eligió en lugar de `rowversion` porque el conflicto es **esperado** (dos cajeros, mismo cliente). Con bloqueo optimista habría que reintentar desde la interfaz; con el bloqueo, el segundo pago simplemente se aplica después.
- **Defensa en profundidad:**
  - El UPDATE exige que el saldo no haya cambiado desde que se leyó.
  - `CHECK` impide saldos negativos o mayores al monto.
  - `UQ_Pago_Referencia` evita aplicar dos veces el mismo comprobante (doble clic o reintento): devuelve −6.
- **Reparto FIFO** en `tvf_PagoDistribucion`. Es la misma función para la vista previa y para la aplicación real, así que lo que el cajero ve es lo que se aplica.
- `READ_COMMITTED_SNAPSHOT` permite que los reportes no bloqueen los pagos ni queden bloqueados por ellos.

## Crystal Reports: pull vs push
- **Estado de cuenta: PULL.** Crystal ejecuta `usp_Reporte_EstadoCuenta` y el subreporte de pagos de 90 días.
  - Conviene cuando el reporte es autocontenido, tiene subreportes que deben consultar con sus propios parámetros, o se va a ejecutar fuera de la aplicación.
  - La conexión se reasigna en tiempo de ejecución desde config.ini (`SetLogOnInfo` en el reporte y en sus subreportes).
- **Antigüedad: PUSH.** VB6 obtiene el Recordset por las capas de negocio y datos y lo entrega con `Database.SetDataSource`.
  - Conviene cuando los datos ya pasaron por reglas o filtros de la aplicación, o cuando el reporte no debe tener credenciales ni acceso a la base.
  - Desventaja: los subreportes y la paginación sobre volúmenes grandes requieren más trabajo.
  - El SP agrega por cliente en el servidor, así que se transfieren unos 2,000 registros en lugar de 100,000.
  - Hallazgo al probar: `SetDataSource` descarta los parámetros del SP usado en el diseño, así que la fecha de corte llegaba vacía. Se resolvió con la fórmula `{@FechaCorteTexto}`, que la app asigna desde el código (`modReportes.AsignarFormulaTexto`).
- Los parámetros se envían desde código con `EnableParameterPrompting = False`. La exportación a PDF usa `ExportOptions` + `Export False`, sin diálogos.
- Fechas `dd/MM/yyyy` y montos con 2 decimales se formatean en el reporte, sin depender de la región.

## Manejo de errores
- `On Error GoTo` en todo procedimiento público.
- Cada error se registra **una sola vez**, donde se origina: archivo `logs\CobranzaExpress_yyyymmdd.log` + tabla `LogError` (módulo, procedimiento, número, descripción, fecha, usuario).
- Luego se vuelve a lanzar con `Source = FUENTE_REGISTRADO` para que las capas superiores solo informen al usuario.
- `On Error Resume Next` solo en dos casos, ambos comentados: el propio registro de errores (no debe provocar otro error) y la liberación de recursos.
- Si la base no responde, el log se escribe en el archivo de todos modos.

## Índices
- `IX_Factura_Cliente_Venc (ClienteId, FechaVencimiento, NumeroFactura) INCLUDE (...)`: seek por cliente, ya ordenado en FIFO. Cubre pendientes, estado de cuenta y el reparto de pagos.
- `IX_Factura_Saldo INCLUDE (ClienteId, FechaVencimiento, Tipo)`: la antigüedad lee solo las columnas necesarias, sin tocar el clustered.
- Evidencia: `sql/07_evidencia_rendimiento.sql` y `docs/evidencia/`.

## Tradeoffs por tiempo y pendientes
- La antigüedad usa el **saldo actual**, no el saldo histórico a la fecha de corte (para eso habría que reconstruir con `PagoDetalle`).
- El saldo a favor se acumula, pero **no se aplica automáticamente** a facturas futuras (requiere un proceso de aplicación de anticipos).
- Las notas de crédito no se cruzan automáticamente contra facturas.
- La clave de `CobranzaApp` está en `04_permisos.sql` como ejemplo. En producción debe generarse aparte y config.ini debería cifrarse.

## Mejoras futuras / modernización por etapas
1. Poner los SP detrás de una API REST (.NET), reutilizando `tvf_Mora` y `usp_Pago_Aplicar` sin cambios.
2. Reemplazar las pantallas VB6 una por una (cliente web o WinForms) que consuman la API, conviviendo con la versión VB6 mientras dura la migración.
3. Migrar los reportes a Crystal para .NET o a SSRS, que leen los **mismos SP**, así que la lógica y los números no cambian.
4. Retirar VB6 cuando todas las pantallas estén migradas.
