VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsDataAccess - UNICO punto del sistema que crea objetos ADO
' (Connection, Command, Recordset). Solo procedimientos almacenados
' con parametros tipados; nunca SQL concatenado.
Option Explicit

Private Const MODULO As String = "clsDataAccess"
Private Const ERR_TIMEOUT As Long = -2147217871        ' 0x80040E31 Query timeout expired
Private mCn As ADODB.Connection

Private Sub Class_Terminate()
    Cerrar
End Sub

Public Sub Cerrar()
    ' On Error Resume Next justificado: liberacion de recursos al terminar;
    ' un fallo al cerrar no debe ocultar el error original.
    On Error Resume Next
    If Not mCn Is Nothing Then
        If mCn.State <> adStateClosed Then mCn.Close
        Set mCn = Nothing
    End If
End Sub

Private Function Conexion() As ADODB.Connection
    If mCn Is Nothing Then
        Set mCn = New ADODB.Connection
        mCn.ConnectionTimeout = modConfig.ConnectionTimeout()
        mCn.CommandTimeout = modConfig.CommandTimeout()
        mCn.CursorLocation = adUseServer
        mCn.Open modConfig.CadenaConexion()
    End If
    Set Conexion = mCn
End Function

' Crea un Command de procedimiento almacenado con su RETURN_VALUE.
Public Function CrearComando(ByVal NombreSP As String) As ADODB.Command
    On Error GoTo EH
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = Conexion()
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = NombreSP
    cmd.CommandTimeout = modConfig.CommandTimeout()
    cmd.Parameters.Append cmd.CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
    Set CrearComando = cmd
    Set cmd = Nothing
    Exit Function
EH:
    Set cmd = Nothing
    ManejarError "CrearComando", NombreSP
End Function

Public Sub AgregarParametro(ByVal cmd As ADODB.Command, ByVal Nombre As String, _
                            ByVal Tipo As ADODB.DataTypeEnum, ByVal Direccion As ADODB.ParameterDirectionEnum, _
                            ByVal Valor As Variant, Optional ByVal Tamano As Long = 0, _
                            Optional ByVal Precision As Byte = 18, Optional ByVal Escala As Byte = 2)
    On Error GoTo EH
    Dim p As ADODB.Parameter
    Set p = cmd.CreateParameter(Nombre, Tipo, Direccion, Tamano)
    If Tipo = adDecimal Or Tipo = adNumeric Then
        p.Precision = Precision
        p.NumericScale = Escala
    End If
    If Direccion = adParamInput Or Direccion = adParamInputOutput Then p.Value = Valor
    cmd.Parameters.Append p
    Set p = Nothing
    Exit Sub
EH:
    Set p = Nothing
    ManejarError "AgregarParametro", Nombre
End Sub

' Recordset DESCONECTADO (cursor del cliente, estatico): para grids.
Public Function AbrirDesconectado(ByVal cmd As ADODB.Command) As ADODB.Recordset
    On Error GoTo EH
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic, adLockBatchOptimistic
    Set rs.ActiveConnection = Nothing
    Set AbrirDesconectado = rs
    Set rs = Nothing
    Exit Function
EH:
    Set rs = Nothing
    ManejarError "AbrirDesconectado", cmd.CommandText
End Function

' Recordset forward-only / read-only (cursor del servidor): para listados
' que se recorren una sola vez. El llamador debe cerrarlo.
Public Function AbrirSoloLectura(ByVal cmd As ADODB.Command) As ADODB.Recordset
    On Error GoTo EH
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
    Set AbrirSoloLectura = rs
    Set rs = Nothing
    Exit Function
EH:
    Set rs = Nothing
    ManejarError "AbrirSoloLectura", cmd.CommandText
End Function

' Ejecuta sin resultados y devuelve el valor de retorno del SP.
Public Function Ejecutar(ByVal cmd As ADODB.Command) As Long
    On Error GoTo EH
    cmd.Execute , , adExecuteNoRecords
    Ejecutar = ValorRetorno(cmd)
    Exit Function
EH:
    ManejarError "Ejecutar", cmd.CommandText
End Function

Public Function ValorRetorno(ByVal cmd As ADODB.Command) As Long
    ValorRetorno = modUtil.NzLng(cmd.Parameters("RETURN_VALUE").Value, 0)
End Function

' Registra el error (incluida la coleccion Errors del proveedor) y lo
' re-lanza con un mensaje claro para el usuario.
Private Sub ManejarError(ByVal Procedimiento As String, ByVal Contexto As String)
    Dim lNum As Long, sDesc As String, sDetalle As String, sMsg As String
    Dim e As ADODB.Error
    lNum = Err.Number
    sDesc = Err.Description

    If Not mCn Is Nothing Then
        For Each e In mCn.Errors
            sDetalle = sDetalle & " | [" & e.NativeError & "/" & e.SQLState & "] " & e.Description
        Next
        mCn.Errors.Clear
    End If

    If lNum = ERR_TIMEOUT Or InStr(1, sDesc & sDetalle, "timeout", vbTextCompare) > 0 _
       Or InStr(1, sDesc & sDetalle, "tiempo de espera", vbTextCompare) > 0 Then
        sMsg = "El servidor de base de datos no respondio a tiempo. Intente de nuevo en unos momentos " & _
               "o contacte a soporte si el problema continua."
    ElseIf mCn Is Nothing Then
        sMsg = "No se pudo preparar la conexion: " & sDesc
    ElseIf mCn.State = adStateClosed Then
        sMsg = "No se pudo conectar al servidor de base de datos. Verifique la red y config.ini." & _
               vbCrLf & "Detalle: " & sDesc
        Set mCn = Nothing
    Else
        sMsg = sDesc
    End If

    modLog.RegistrarError MODULO, Procedimiento & " (" & Contexto & ")", lNum, sDesc & sDetalle
    Err.Raise lNum, modLog.FUENTE_REGISTRADO, sMsg
End Sub
