Attribute VB_Name = "modLog"
' modLog - registro centralizado de errores
' Escribe en archivo (siempre) y en la tabla LogError (via SP).
' Convencion: un error ya registrado se re-lanza con Source = FUENTE_REGISTRADO
' para que las capas superiores no lo registren dos veces.
Option Explicit

Public Const FUENTE_REGISTRADO As String = "CobranzaExpress.Registrado"
Private mRegistrando As Boolean   ' evita recursion si la BD tambien falla

Public Function UsuarioActual() As String
    UsuarioActual = Environ$("USERNAME")
End Function

Public Function EquipoActual() As String
    EquipoActual = Environ$("COMPUTERNAME")
End Function

Public Sub RegistrarError(ByVal Modulo As String, ByVal Procedimiento As String, _
                          ByVal Numero As Long, ByVal Descripcion As String)
    ' On Error Resume Next justificado: el registro de errores nunca debe
    ' provocar un nuevo error ni ocultar el error original al usuario.
    On Error Resume Next
    EscribirArchivo Modulo, Procedimiento, Numero, Descripcion

    If Not mRegistrando Then
        mRegistrando = True
        Dim da As clsDataAccess
        Dim cmd As ADODB.Command
        Set da = New clsDataAccess
        Set cmd = da.CrearComando("dbo.usp_LogError_Insertar")
        If Not cmd Is Nothing Then
            da.AgregarParametro cmd, "@Usuario", adVarWChar, adParamInput, UsuarioActual(), 128
            da.AgregarParametro cmd, "@Equipo", adVarWChar, adParamInput, EquipoActual(), 128
            da.AgregarParametro cmd, "@Modulo", adVarWChar, adParamInput, Left$(Modulo, 100), 100
            da.AgregarParametro cmd, "@Procedimiento", adVarWChar, adParamInput, Left$(Procedimiento, 100), 100
            da.AgregarParametro cmd, "@Numero", adInteger, adParamInput, Numero
            da.AgregarParametro cmd, "@Descripcion", adVarWChar, adParamInput, Left$(Descripcion, 2000), 2000
            da.Ejecutar cmd
        End If
        Set cmd = Nothing
        da.Cerrar
        Set da = Nothing
        mRegistrando = False
    End If
End Sub

Private Sub EscribirArchivo(ByVal Modulo As String, ByVal Procedimiento As String, _
                            ByVal Numero As Long, ByVal Descripcion As String)
    On Error Resume Next   ' justificado: ver RegistrarError
    Dim sRuta As String
    Dim iArch As Integer
    sRuta = modConfig.RutaLog()
    If Len(Dir$(sRuta, vbDirectory)) = 0 Then MkDir sRuta
    iArch = FreeFile
    Open sRuta & "\CobranzaExpress_" & Format$(Date, "yyyymmdd") & ".log" For Append As #iArch
    Print #iArch, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & UsuarioActual() & vbTab & _
                  Modulo & "." & Procedimiento & vbTab & Numero & vbTab & _
                  Replace(Replace(Descripcion, vbCr, " "), vbLf, " ")
    Close #iArch
End Sub

' Para formularios: registra (si no estaba registrado) y avisa al usuario.
Public Sub MostrarError(ByVal Modulo As String, ByVal Procedimiento As String, _
                        ByVal Numero As Long, ByVal Descripcion As String, ByVal Fuente As String)
    If Fuente <> FUENTE_REGISTRADO Then RegistrarError Modulo, Procedimiento, Numero, Descripcion
    MsgBox Descripcion, vbExclamation, "Cobranza Express"
End Sub
