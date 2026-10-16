# Cobranza Express

Módulo de cuentas por cobrar de escritorio, hecho en **Visual Basic 6 + ADO + SQL Server + Crystal Reports**:
- Consulta de clientes con saldo e interés moratorio.
- Registro de pagos con aplicación FIFO atómica.
- Reportes de estado de cuenta y de antigüedad de saldos, con exportación a PDF.

## Requisitos y versiones usadas
| Componente | Versión |
|---|---|
| Windows | 11 Enterprise (VM Hyper-V) |
| Visual Basic | 6.0 Enterprise **SP6** |
| SQL Server | 2019 Express (`.\SQLEXPRESS`), intercalación `Modern_Spanish_CI_AS` |
| Crystal Reports | **XI Release 2** (11.5), RDC `craxdrt.dll` + `CRViewer.dll` |
| ADO | Microsoft ActiveX Data Objects **2.8** Library |
| Proveedor OLE DB | `SQLOLEDB` (incluido en Windows) |
| Controles | Microsoft FlexGrid 6.0 (`MSFLXGRD.OCX`) |

Referencias del proyecto: OLE Automation, Microsoft ActiveX Data Objects 2.8, Crystal Reports ActiveX Designer Run Time Library 11.5 y Crystal ActiveX Report Viewer Library 11.5.

## Estructura
```
sql/       scripts de base de datos (ejecutar en orden)
src/       proyecto VB6 (CobranzaExpress.vbp)
reportes/  EstadoCuenta.rpt, AntiguedadSaldos.rpt
pdf/       PDF de ejemplo exportados
docs/      evidencia de rendimiento, requisitos y riesgos
```

## 1. Base de datos
En SSMS, con un usuario administrador, ejecuta los scripts en orden:
1. `sql/01_crear_bd.sql`: crea la base `Cobranza`.
2. `sql/02_estructura.sql`: tablas, llaves, CHECK e índices.
3. `sql/03_procedimientos.sql`: funciones y procedimientos almacenados.
4. `sql/04_permisos.sql`: usuario `CobranzaApp`, solo EXECUTE. **Cambia la clave antes de ejecutarlo.**
5. `sql/05_datos_iniciales.sql`: 8 clientes, uno por escenario (sin saldo, al día, cada rango de mora, nota de crédito, datos incompletos, FIFO).
6. `sql/06_datos_volumen.sql`: 2,000 clientes y 100,000 facturas.
7. `sql/07_evidencia_rendimiento.sql`: estadísticas y plan de ejecución (opcional).

## 2. Configuración
Copia `src/config.ini.example` como `src/config.ini` (en la misma carpeta del .exe) y completa `Clave=` con la clave de `CobranzaApp`. `config.ini` no se versiona.

## 3. Compilar y ejecutar
1. Abre `src/CobranzaExpress.vbp` en VB6. En Windows 10/11 hay que ejecutar VB6 como administrador.
2. Para correr: **Ctrl+F5**. Para compilar: **File → Make CobranzaExpress.exe**.
3. El ejecutable compilado está en la sección *Releases* del repositorio.

## 4. Uso
- **Pruebas de interés moratorio**: ejecuta 12 casos y compara esperado vs. VB6 vs. SQL.
- **Consulta de clientes**: búsqueda parcial por código, nombre o NIT. Al seleccionar un cliente se muestran sus facturas en orden FIFO, con el interés calculado en VB6 y en SQL.
- **Registro de pago**: filtra el cliente, captura los datos, genera la vista previa y aplica el pago.
- **Reportes**: elige el reporte y sus parámetros, y luego verlo en pantalla o exportarlo a PDF (`pdf/`).

## Documentación
- [DECISIONES.md](DECISIONES.md): arquitectura, ADO, dinero, mora, concurrencia, pull/push, errores y pendientes.
- [REVISION.md](REVISION.md): revisión del código legado.
- [docs/REQUISITOS_Y_RIESGOS.md](docs/REQUISITOS_Y_RIESGOS.md): requisitos funcionales y no funcionales, riesgos y puesta en marcha.
