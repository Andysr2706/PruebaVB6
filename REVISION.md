# Revisión de código legado: `cmdAplicar_Click`

Se revisó el procedimiento como en una revisión de código del equipo. Hay **16 problemas**. Cada uno indica su tipo, su severidad y cómo queda resuelto en Cobranza Express.

## 1. Problemas encontrados

| # | Problema | Tipo | Severidad | Resuelto en Cobranza Express |
|---|---|---|---|---|
| 1 | **El INSERT en `Pago` usa `monto` *después* de modificarlo.** Cuando la factura se paga completa, registra lo que *sobra* del pago. Cuando el pago es parcial, registra **0** (porque `monto = 0` se asignó antes del INSERT). El monto guardado nunca es el monto aplicado. | Lógica de negocio | **Alta** | `usp_Pago_Aplicar` calcula `MontoAplicado` por factura y lo guarda en `PagoDetalle` (`CK_PagoDetalle_Cuadre`: SaldoNuevo = SaldoAnterior − MontoAplicado). |
| 2 | **No hay transacción.** Cada factura se actualiza y su pago se inserta por separado. Si algo falla a la mitad, quedan facturas rebajadas sin su pago, o al revés. | Integridad de datos | **Alta** | `usp_Pago_Aplicar`: `BEGIN TRAN` + `TRY/CATCH` + `XACT_ABORT ON`. O se aplica todo o no se aplica nada. |
| 3 | **`On Error Resume Next` en todo el procedimiento**, y el mensaje "Pago aplicado correctamente" se muestra siempre. Un error de conexión, de conversión o de bloqueo se ignora y el usuario cree que el pago entró. | Integridad / Mantenibilidad | **Alta** | `On Error GoTo` en todos los procedimientos públicos. Registro en archivo y en `LogError` (`modLog`). El resultado se toma del valor de retorno del SP. |
| 4 | **Inyección SQL**: `txtCliente.Text` se concatena en el SELECT. Con `1 OR 1=1` se aplicaría el pago a facturas de todos los clientes, y otros valores permiten borrar datos. | Seguridad | **Alta** | Todo pasa por `ADODB.Command` con `CreateParameter` tipado (`clsDataAccess.AgregarParametro`). No hay SQL concatenado. |
| 5 | **Credenciales `sa` / `Admin123` escritas en el código** (quedan en el .exe y en el control de versiones), con servidor fijo. | Seguridad | **Alta** | Conexión en `config.ini` (fuera de Git). Usuario `CobranzaApp` con **solo EXECUTE** (`04_permisos.sql`). |
| 6 | **Concurrencia**: lee el saldo, lo resta en el cliente y escribe el nuevo valor (cursor dinámico optimista). Dos cajeros con el mismo cliente leen el mismo saldo y aplican dos veces. Si el bloqueo optimista falla, `Resume Next` lo oculta. | Integridad de datos | **Alta** | `UPDLOCK, HOLDLOCK` sobre la fila del cliente (serializa los pagos del mismo cliente). El UPDATE verifica que el saldo no cambió. Los CHECK impiden saldos negativos. |
| 7 | **Se pierde el excedente**: si el pago supera la deuda, lo que sobra no se registra en ningún lado. | Lógica de negocio | **Alta** | El excedente se guarda en `Pago.SaldoAFavor` y se acumula en `Cliente.SaldoAFavor`. `CK_Pago_Cuadre`: MontoAplicado + SaldoAFavor = Monto. |
| 8 | **Sin orden de aplicación**: el SELECT no tiene `ORDER BY`, así que el pago cae en facturas en orden arbitrario, no FIFO. | Lógica de negocio | **Alta** | `tvf_PagoDistribucion`: vencimiento ascendente, empate por número de factura, las facturas sin vencimiento al final. |
| 9 | **No valida el monto**: `CDbl(txtMonto.Text)` acepta 0 y negativos (un monto negativo **aumenta** los saldos) y falla con texto vacío. Además depende de la configuración regional ("1,500" puede leerse como 1.5). | Lógica / Integridad | **Alta** | `modUtil.ParseMonto` (independiente de la región, máximo 2 decimales). Validación en `frmPago`, `clsCobranza.ValidarPago` y en el SP (retorno −1). `CK_Pago_Monto`. |
| 10 | **`Double` para dinero**: errores de punto flotante (0.1 + 0.2 ≠ 0.3). Una resta puede dejar saldos como 0.0000001 que nunca llegan a 0. | Integridad de datos | Media | `DECIMAL(18,2)` en SQL, `Currency` y `Decimal` en VB6. |
| 11 | **Fecha concatenada con `Date`**: el texto depende de la región (`16/10/2026` puede leerse como mes 16 y fallar, o invertir día y mes). Con `Resume Next` el INSERT falla en silencio. | Integridad de datos | Media | Parámetros `adDBTimeStamp`. Captura `dd/mm/aaaa` validada con `ParseFecha`. |
| 12 | **El bucle sigue después de `monto = 0`**: inserta un "pago" de 0 por cada factura restante y hace UPDATE innecesarios. | Lógica / Rendimiento | Media | Solo se procesan las facturas con `MontoAplicado > 0`. `CK_PagoDetalle_Monto > 0`. |
| 13 | **Modelo de datos sin trazabilidad**: `Pago` se liga a una factura, sin cliente, forma de pago, referencia ni usuario. No se puede reconstruir qué pagó quién, ni detectar duplicados. | Mantenibilidad / Integridad | Media | `Pago` (encabezado con cliente, forma, referencia, usuario) + `PagoDetalle`. `UQ_Pago_Referencia` evita aplicar dos veces el mismo comprobante. |
| 14 | **Cursor dinámico del servidor + `SELECT *`** para leer y actualizar fila por fila: mantiene bloqueos y hace un viaje de red por cada fila. | Rendimiento | Media | Todo el proceso se hace de una vez dentro del SP, con lecturas set-based. Cursores forward-only y desconectados según el uso. |
| 15 | **Lógica de negocio y SQL dentro del formulario**: no se puede probar ni reutilizar, y la regla se duplica en cada pantalla. | Mantenibilidad | Media | Capas: formularios → `clsCobranza` → `clsDataAccess` → SP. |
| 16 | **`Dim ... As New` y objetos sin cerrar** (`cn`, `rs` nunca se cierran ni se liberan): conexiones abiertas en el servidor. | Rendimiento / Mantenibilidad | Baja | `Set x = New ...`, `Close` y `Set Nothing` explícitos. `clsDataAccess.Cerrar` en `Class_Terminate`. |

## 2. ¿Qué causa los descuadres reportados?

Los descuadres vienen principalmente de **#1, #7, #2/#3 y #6**:

- **Escenario A (#1, el más frecuente).** El cliente debe F1 = 1,000 y F2 = 500 y paga 1,200.
  - Primera vuelta: `monto` pasa de 1,200 a 200 y F1 queda en 0, pero se inserta un Pago de **200**.
  - Segunda vuelta: F2 queda en 300, `monto` pasa a 0 y se inserta un Pago de **0**.
  - Resultado: las facturas bajaron **1,200**, pero en la tabla `Pago` hay **200**. Facturas y pagos no cuadran.
- **Escenario B (#7).** El cliente debe 500 y paga 800. Los 300 de excedente no se registran: entró dinero a caja que el sistema no tiene.
- **Escenario C (#2 + #3).** Se cae la red o hay un bloqueo en la segunda factura. La primera ya quedó rebajada, la segunda no, y el usuario ve "Pago aplicado correctamente" y lo vuelve a intentar: el pago se aplica dos veces.
- **Escenario D (#6).** Dos cajeros cobran al mismo cliente al mismo tiempo. Ambos leen saldo 1,000 y cada uno escribe su resta, así que una actualización pisa a la otra, o el conflicto se ignora por `Resume Next`. Ambos pagos quedan insertados pero solo uno se refleja en el saldo.
- **Escenario E (#9).** Un monto negativo o mal interpretado por la región ("1,500" leído como 1.5) aumenta saldos o aplica cantidades equivocadas.

## 3. Verificación en la implementación

- Doble aplicación: aplicar dos veces el mismo pago (misma forma y referencia) devuelve el código −6 y no cambia nada (probado en `frmPago`).
- Cuadre: `SELECT` de `Pago` vs. la suma de `PagoDetalle` + `SaldoAFavor`. El CHECK lo impide a nivel de base de datos.
