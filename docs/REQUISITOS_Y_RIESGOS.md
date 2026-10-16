# Requisitos, riesgos y puesta en marcha

## Requisitos funcionales
| ID | Requisito |
|---|---|
| RF-01 | Buscar clientes por código, nombre o NIT (búsqueda parcial). |
| RF-02 | Mostrar por cliente: límite, saldo total, saldo vencido, días de mayor atraso, interés moratorio y saldo a favor. |
| RF-03 | Listar las facturas pendientes del cliente ordenadas por vencimiento (FIFO). |
| RF-04 | Registrar un pago con cliente, fecha, monto, forma de pago y referencia. |
| RF-05 | Mostrar una vista previa de la aplicación del pago antes de confirmar. |
| RF-06 | Aplicar el pago de forma atómica en orden FIFO. El excedente se registra como saldo a favor. |
| RF-07 | Impedir pagos duplicados (misma forma de pago y referencia para el mismo cliente). |
| RF-08 | Calcular el interés moratorio con días de gracia, tasa del cliente o tasa por defecto, y redondeo comercial. |
| RF-09 | Marcar las facturas con datos incompletos (sin vencimiento, sin tasa). |
| RF-10 | Reporte de estado de cuenta con un subreporte de los pagos de los últimos 90 días. |
| RF-11 | Reporte de antigüedad de saldos agrupado por vendedor y cliente, con subtotales y total general. |
| RF-12 | Exportar los reportes a PDF sin intervención del usuario. |

## Requisitos no funcionales
| ID | Requisito |
|---|---|
| RNF-01 | Integridad: ningún pago deja cambios parciales y ningún saldo queda negativo (transacción + CHECK). |
| RNF-02 | Concurrencia: varios cajeros pueden cobrar al mismo cliente al mismo tiempo sin aplicaciones duplicadas. |
| RNF-03 | Seguridad: sin credenciales en el código; usuario de BD con solo EXECUTE; sin SQL concatenado. |
| RNF-04 | Consistencia: el interés es idéntico en VB6, SQL Server y Crystal. |
| RNF-05 | Rendimiento: el reporte de antigüedad responde en segundos con más de 100,000 facturas. |
| RNF-06 | Trazabilidad: errores registrados en archivo y en `LogError`; pagos con usuario y fecha de registro. |
| RNF-07 | Independencia regional: fechas `dd/mm/aaaa` y montos sin depender de la configuración de Windows. |
| RNF-08 | Mantenibilidad: capas separadas y configuración externa (config.ini). |

## Riesgos
| Riesgo | Impacto | Mitigación |
|---|---|---|
| VB6 y Crystal XI R2 sin soporte del fabricante | Alto | Mantener la lógica en los SP (reutilizable) y seguir la ruta de modernización por etapas (DECISIONES.md). |
| Instalar el runtime de Crystal y los OCX en las estaciones | Medio | Instalador con el merge module / runtime de Crystal XI R2, MSFLXGRD.OCX y CRViewer; probar en una estación limpia. |
| Clave en config.ini en texto plano | Medio | Permisos NTFS sobre la carpeta, cifrado de la sección o autenticación integrada de Windows. |
| La antigüedad usa el saldo actual (no el histórico a la fecha de corte) | Medio | Documentado; reconstruir el saldo desde `PagoDetalle` si se requiere un corte histórico. |
| Datos migrados incompletos (sin vencimiento o sin tasa) | Medio | Se marcan en pantalla y en el estado de cuenta; depurar antes de la puesta en marcha. |
| Saldo a favor no se aplica automáticamente | Bajo | Proceso manual o aplicación de anticipos en una siguiente versión. |

## Acciones para la puesta en marcha
1. Crear la base y ejecutar los scripts 01–04 en el servidor de producción, con una clave segura para `CobranzaApp`.
2. Migrar clientes y facturas reales, validar los saldos contra el sistema anterior y depurar los datos incompletos.
3. Definir `TASA_MORA_DEFECTO` y las tasas y días de gracia por cliente con el área de créditos.
4. Instalar en las estaciones: el ejecutable, config.ini, el runtime de Crystal XI R2, MSFLXGRD.OCX y CRViewer.
5. Prueba piloto con 2 cajeros cobrando al mismo cliente (concurrencia) y cuadre diario de `Pago` contra caja.
6. Plan de respaldo de la base y monitoreo de la tabla `LogError`.
