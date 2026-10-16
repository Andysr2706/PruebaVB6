# Evidencia de rendimiento

- **Volumen:** 100,000 facturas de volumen + datos iniciales, 2,008 clientes y 10 vendedores (`sql/06_datos_volumen.sql`).
- **Script:** `sql/07_evidencia_rendimiento.sql` con `SET STATISTICS IO, TIME ON` y plan de ejecución real.
- **Entorno:** SQL Server 2019 Express en una VM con 4 vCPU y 8 GB.

| # | Consulta | Lecturas lógicas `Factura` | CPU | Tiempo |
|---|---|---|---|---|
| 1 | `usp_Reporte_Antiguedad` (2,007 clientes) | 829 | 109–125 ms | 240–380 ms |
| 2 | Agregado equivalente forzando el índice clustered | 859 | 63 ms | 137 ms |
| 3 | `usp_Reporte_EstadoCuenta` (1 cliente) | **3** | 15 ms | 28 ms |
| 4 | Las mismas 34 facturas del cliente forzando scan (sin índice) | **859** | 0 ms | 142 ms |

## Lectura de los resultados
- **Antigüedad de saldos.** El reporte necesita *todas* las facturas con saldo, unas 2/3 del total, así que un recorrido completo es inevitable. El costo real está en la agregación por cliente.
  - **El plan real (`plan_antiguedad.png`) mostró algo inesperado:** el optimizador **no usó `IX_Factura_Saldo`**. Prefirió recorrer `IX_Factura_Cliente_Venc`, que ya viene ordenado por `ClienteId` y le permite un *Merge Join* con `Cliente` sin ordenar antes.
  - **Decisión: se eliminó `IX_Factura_Saldo`.** No mejoraba la consulta y encarecía cada `UPDATE` de saldo en `usp_Pago_Aplicar`.
  - **Índice sugerido por SQL Server `(FechaEmision, Saldo) INCLUDE (ClienteId, FechaVencimiento)`, impacto estimado 29%: descartado.** El filtro `FechaEmision <= @FechaCorte` incluye casi todas las filas, así que el índice sería otra copia completa de la tabla con más costo de escritura.
  - El ⚠️ en *Hash Match (Aggregate)* es un *spill* menor a tempdb (`Workfile`, 24 lecturas). No es relevante con 240 ms.
  - El resultado es aceptable: menos de 0.3 s para 100k facturas. Además, solo viajan unos 2,000 registros a la estación (modelo push) en lugar de 100,000.
  - Con millones de filas, las opciones serían un índice columnstore no agrupado sobre `Factura`, o una tabla de saldos resumidos por cliente que se actualice en `usp_Pago_Aplicar`.
- **Estado de cuenta, pendientes y reparto de pagos.** Aquí `IX_Factura_Cliente_Venc` sí es decisivo: el seek por `ClienteId` lee **3 páginas**; sin el índice, para las mismas 34 facturas se leen **859 páginas** (≈286 veces más), es decir, toda la tabla (consulta 4). Este costo crece con el volumen; el seek no. Además, el índice ya entrega las filas en orden FIFO (vencimiento, número), así que no hace falta ordenar.
- `READ_COMMITTED_SNAPSHOT` evita que los reportes bloqueen la aplicación de pagos.

## Archivos
- `messages.txt`: salida completa de STATISTICS IO/TIME.
- `plan_antiguedad.png` / `.sqlplan`: plan de ejecución real.
