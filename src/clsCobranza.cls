VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCobranza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================================
' clsCobranza - reglas de negocio de cuentas por cobrar.
' Valida entradas, invoca la capa de datos y traduce resultados.
' No conoce formularios ni crea objetos ADO de conexion.
'=====================================================================
Option Explicit

Private Const MODULO As String = "clsCobranza"
Private Const ERR_VALIDACION As Long = vbObjectError + 2000

Public Enum ResultadoPago
    rpOK = 0
    rpMontoInvalido = -1
    rpFechaInvalida = -2
    rpFormaPagoInvalida = -3
    rpReferenciaVacia = -4
    rpClienteNoExiste = -5
    rpPagoDuplicado = -6
End Enum

Private mDA As clsDataAccess

Private Sub Class_Initialize()
    Set mDA = New clsDataAccess
End Sub

Private Sub Class_Terminate()
    If Not mDA Is Nothing Then mDA.Cerrar
    Set mDA = Nothing
End Sub

'---------------------------------------------------------------------
' Consulta de clientes con saldo (Recordset desconectado para grid)
'---------------------------------------------------------------------
Public Function BuscarClientes(ByVal Texto As String, ByVal FechaCorte As Date) As ADODB.Recordset
    On Error GoTo EH
    Dim cmd As ADODB.Command
    Set cmd = mDA.CrearComando("dbo.usp_Cliente_Buscar")
    mDA.AgregarParametro cmd, "@Texto", adVarWChar, adParamInput, IIf(Len(Trim$(Texto)) = 0, Null, Trim$(Texto)), 100
    mDA.AgregarParametro cmd, "@FechaCorte", adDBTimeStamp, adParamInput, FechaCorte
    Set BuscarClientes = mDA.AbrirDesconectado(cmd)
    Set cmd = Nothing
    Exit Function
EH:
    Set cmd = Nothing
    Relanzar "BuscarClientes"
End Function

Public Function FacturasPendientes(ByVal ClienteId As Long, ByVal FechaCorte As Date) As ADODB.Recordset
    On Error GoTo EH
    Dim cmd As ADODB.Command
    Set cmd = mDA.CrearComando("dbo.usp_Factura_PendientesPorCliente")
    mDA.AgregarParametro cmd, "@ClienteId", adInteger, adParamInput, ClienteId
    mDA.AgregarParametro cmd, "@FechaCorte", adDBTimeStamp, adParamInput, FechaCorte
    Set FacturasPendientes = mDA.AbrirDesconectado(cmd)
    Set cmd = Nothing
    Exit Function
EH:
    Set cmd = Nothing
    Relanzar "FacturasPendientes"
End Function

' Lista para combos: recorre un cursor forward-only/read-only y devuelve
' una coleccion de arreglos (ClienteId, "Codigo - Nombre").
Public Function ListaClientes() As Collection
    On Error GoTo EH
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim col As Collection
    Set col = New Collection
    Set cmd = mDA.CrearComando("dbo.usp_Cliente_Buscar")
    mDA.AgregarParametro cmd, "@Texto", adVarWChar, adParamInput, Null, 100
    mDA.AgregarParametro cmd, "@FechaCorte", adDBTimeStamp, adParamInput, Date
    Set rs = mDA.AbrirSoloLectura(cmd)
    Do While Not rs.EOF
        col.Add Array(modUtil.NzLng(rs!ClienteId), _
                      modUtil.NzStr(rs!Codigo) & " - " & modUtil.NzStr(rs!Nombre) & _
                      " (" & modUtil.NzStr(rs!NIT, "sin NIT") & ")")
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    Set ListaClientes = col
    Exit Function
EH:
    If Not rs Is Nothing Then
        If rs.State <> adStateClosed Then rs.Close
    End If
    Set rs = Nothing
    Set cmd = Nothing
    Relanzar "ListaClientes"
End Function

Public Function TasaDefecto() As Variant
    On Error GoTo EH
    Dim cmd As ADODB.Command
    Set cmd = mDA.CrearComando("dbo.usp_Parametro_TasaDefecto")
    mDA.AgregarParametro cmd, "@Tasa", adDecimal, adParamOutput, Null, , 9, 6
    mDA.Ejecutar cmd
    TasaDefecto = cmd.Parameters("@Tasa").Value     ' puede ser Null
    Set cmd = Nothing
    Exit Function
EH:
    Set cmd = Nothing
    Relanzar "TasaDefecto"
End Function

'---------------------------------------------------------------------
' Pagos
'---------------------------------------------------------------------
Public Sub ValidarPago(ByVal ClienteId As Long, ByVal Fecha As Date, ByVal Monto As Currency, _
                       ByVal FormaPago As String, ByVal Referencia As String)
    If ClienteId <= 0 Then Err.Raise ERR_VALIDACION, MODULO, "Seleccione un cliente."
    If Monto <= 0 Then Err.Raise ERR_VALIDACION, MODULO, "El monto debe ser mayor que cero."
    If Fecha > Date Then Err.Raise ERR_VALIDACION, MODULO, "La fecha del pago no puede ser futura."
    If Len(Trim$(FormaPago)) = 0 Then Err.Raise ERR_VALIDACION, MODULO, "Seleccione la forma de pago."
    If Len(Trim$(Referencia)) = 0 Then Err.Raise ERR_VALIDACION, MODULO, "Ingrese la referencia del pago."
End Sub

' Vista previa: misma funcion FIFO que usa la aplicacion real (tvf_PagoDistribucion)
Public Function SimularPago(ByVal ClienteId As Long, ByVal Monto As Currency, _
                            ByRef SaldoAFavor As Currency) As ADODB.Recordset
    On Error GoTo EH
    Dim cmd As ADODB.Command
    If Monto <= 0 Then Err.Raise ERR_VALIDACION, MODULO, "El monto debe ser mayor que cero."
    Set cmd = mDA.CrearComando("dbo.usp_Pago_Simular")
    mDA.AgregarParametro cmd, "@ClienteId", adInteger, adParamInput, ClienteId
    mDA.AgregarParametro cmd, "@Monto", adDecimal, adParamInput, CDec(Monto), , 18, 2
    mDA.AgregarParametro cmd, "@SaldoAFavor", adDecimal, adParamOutput, Null, , 18, 2
    Set SimularPago = mDA.AbrirDesconectado(cmd)
    SaldoAFavor = modUtil.NzCur(cmd.Parameters("@SaldoAFavor").Value)
    Set cmd = Nothing
    Exit Function
EH:
    Set cmd = Nothing
    Relanzar "SimularPago"
End Function

' Aplicacion atomica (toda la transaccion vive en usp_Pago_Aplicar).
' Devuelve el codigo de resultado; PagoId y SaldoAFavor por referencia.
Public Function AplicarPago(ByVal ClienteId As Long, ByVal Fecha As Date, ByVal Monto As Currency, _
                            ByVal FormaPago As String, ByVal Referencia As String, _
                            ByRef PagoId As Long, ByRef SaldoAFavor As Currency) As ResultadoPago
    On Error GoTo EH
    Dim cmd As ADODB.Command
    ValidarPago ClienteId, Fecha, Monto, FormaPago, Referencia

    Set cmd = mDA.CrearComando("dbo.usp_Pago_Aplicar")
    mDA.AgregarParametro cmd, "@ClienteId", adInteger, adParamInput, ClienteId
    mDA.AgregarParametro cmd, "@Fecha", adDBTimeStamp, adParamInput, Fecha
    mDA.AgregarParametro cmd, "@Monto", adDecimal, adParamInput, CDec(Monto), , 18, 2
    mDA.AgregarParametro cmd, "@FormaPago", adVarChar, adParamInput, FormaPago, 15
    mDA.AgregarParametro cmd, "@Referencia", adVarChar, adParamInput, Trim$(Referencia), 50
    mDA.AgregarParametro cmd, "@Usuario", adVarWChar, adParamInput, modLog.UsuarioActual(), 128
    mDA.AgregarParametro cmd, "@PagoId", adInteger, adParamOutput, Null
    mDA.AgregarParametro cmd, "@SaldoAFavor", adDecimal, adParamOutput, Null, , 18, 2

    AplicarPago = mDA.Ejecutar(cmd)
    PagoId = modUtil.NzLng(cmd.Parameters("@PagoId").Value)
    SaldoAFavor = modUtil.NzCur(cmd.Parameters("@SaldoAFavor").Value)
    Set cmd = Nothing
    Exit Function
EH:
    Set cmd = Nothing
    Relanzar "AplicarPago"
End Function

Public Function MensajeResultado(ByVal Codigo As ResultadoPago) As String
    Select Case Codigo
        Case rpOK:                MensajeResultado = "Pago aplicado correctamente."
        Case rpMontoInvalido:     MensajeResultado = "El monto del pago no es valido."
        Case rpFechaInvalida:     MensajeResultado = "La fecha del pago no es valida."
        Case rpFormaPagoInvalida: MensajeResultado = "La forma de pago no es valida."
        Case rpReferenciaVacia:   MensajeResultado = "La referencia es obligatoria."
        Case rpClienteNoExiste:   MensajeResultado = "El cliente no existe."
        Case rpPagoDuplicado:     MensajeResultado = "Ya existe un pago con esa forma de pago y referencia para este cliente. No se aplico de nuevo."
        Case Else:                MensajeResultado = "Resultado desconocido (" & Codigo & ")."
    End Select
End Function

'---------------------------------------------------------------------
' Reportes y pruebas
'---------------------------------------------------------------------
' Modelo PUSH: el Recordset se entrega al reporte de antiguedad
Public Function DatosAntiguedad(ByVal FechaCorte As Date) As ADODB.Recordset
    On Error GoTo EH
    Dim cmd As ADODB.Command
    Set cmd = mDA.CrearComando("dbo.usp_Reporte_Antiguedad")
    mDA.AgregarParametro cmd, "@FechaCorte", adDBTimeStamp, adParamInput, FechaCorte
    Set DatosAntiguedad = mDA.AbrirDesconectado(cmd)
    Set cmd = Nothing
    Exit Function
EH:
    Set cmd = Nothing
    Relanzar "DatosAntiguedad"
End Function

' Calculo de mora en SQL (tvf_Mora) para comparar con clsCalculoMora
Public Function MoraSQL(ByVal Saldo As Currency, ByVal Tasa As Variant, ByVal DiasGracia As Variant, _
                        ByVal FechaVenc As Variant, ByVal FechaCorte As Date, _
                        ByRef DiasMora As Long, ByRef Senal As String) As Currency
    On Error GoTo EH
    Dim cmd As ADODB.Command
    Set cmd = mDA.CrearComando("dbo.usp_Mora_Calcular")
    mDA.AgregarParametro cmd, "@Saldo", adDecimal, adParamInput, CDec(Saldo), , 18, 2
    mDA.AgregarParametro cmd, "@Tasa", adDecimal, adParamInput, Tasa, , 9, 6
    mDA.AgregarParametro cmd, "@DiasGracia", adInteger, adParamInput, DiasGracia
    mDA.AgregarParametro cmd, "@FechaVenc", adDBTimeStamp, adParamInput, FechaVenc
    mDA.AgregarParametro cmd, "@FechaCorte", adDBTimeStamp, adParamInput, FechaCorte
    mDA.AgregarParametro cmd, "@DiasMora", adInteger, adParamOutput, Null
    mDA.AgregarParametro cmd, "@Interes", adDecimal, adParamOutput, Null, , 18, 2
    mDA.AgregarParametro cmd, "@Senal", adVarChar, adParamOutput, Null, 20
    mDA.Ejecutar cmd
    DiasMora = modUtil.NzLng(cmd.Parameters("@DiasMora").Value)
    Senal = modUtil.NzStr(cmd.Parameters("@Senal").Value)
    MoraSQL = modUtil.NzCur(cmd.Parameters("@Interes").Value)
    Set cmd = Nothing
    Exit Function
EH:
    Set cmd = Nothing
    Relanzar "MoraSQL"
End Function

' Registra (si no lo estaba) y propaga el error a la capa superior.
Private Sub Relanzar(ByVal Procedimiento As String)
    Dim lNum As Long, sDesc As String, sFuente As String
    lNum = Err.Number: sDesc = Err.Description: sFuente = Err.Source
    If lNum = ERR_VALIDACION Then
        Err.Raise lNum, modLog.FUENTE_REGISTRADO, sDesc    ' validacion: no es un fallo del sistema
    ElseIf sFuente <> modLog.FUENTE_REGISTRADO Then
        modLog.RegistrarError MODULO, Procedimiento, lNum, sDesc
    End If
    Err.Raise lNum, modLog.FUENTE_REGISTRADO, sDesc
End Sub
