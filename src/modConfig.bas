Attribute VB_Name = "modConfig"
' modConfig - lectura de configuracion externa (config.ini)
' Ningun servidor, usuario o clave vive en el codigo.
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
     ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Const MODULO As String = "modConfig"

Public Function RutaIni() As String
    RutaIni = App.Path & "\config.ini"
End Function

Public Function ConfigExiste() As Boolean
    On Error GoTo EH
    ConfigExiste = (Len(Dir$(RutaIni())) > 0)
    Exit Function
EH:
    ConfigExiste = False
End Function

Public Function LeerIni(ByVal Seccion As String, ByVal Clave As String, Optional ByVal Defecto As String = "") As String
    On Error GoTo EH
    Dim sBuf As String
    Dim lLen As Long
    sBuf = String$(512, vbNullChar)
    lLen = GetPrivateProfileString(Seccion, Clave, Defecto, sBuf, Len(sBuf), RutaIni())
    LeerIni = Trim$(Left$(sBuf, lLen))
    Exit Function
EH:
    Err.Raise Err.Number, MODULO & ".LeerIni", Err.Description
End Function

Public Function CadenaConexion() As String
    On Error GoTo EH
    If Not ConfigExiste() Then
        Err.Raise vbObjectError + 1000, MODULO, "No se encontro el archivo de configuracion: " & RutaIni()
    End If
    CadenaConexion = "Provider=" & LeerIni("BaseDatos", "Proveedor", "SQLOLEDB") & ";" & _
                     "Data Source=" & LeerIni("BaseDatos", "Servidor") & ";" & _
                     "Initial Catalog=" & LeerIni("BaseDatos", "BaseDatos", "Cobranza") & ";" & _
                     "User ID=" & LeerIni("BaseDatos", "Usuario") & ";" & _
                     "Password=" & LeerIni("BaseDatos", "Clave") & ";" & _
                     "Application Name=CobranzaExpress;"
    Exit Function
EH:
    Err.Raise Err.Number, MODULO & ".CadenaConexion", Err.Description
End Function

' Datos de conexion para reasignar el origen de los reportes Crystal
Public Function Servidor() As String
    Servidor = LeerIni("BaseDatos", "Servidor")
End Function

Public Function BaseDatos() As String
    BaseDatos = LeerIni("BaseDatos", "BaseDatos", "Cobranza")
End Function

Public Function UsuarioBD() As String
    UsuarioBD = LeerIni("BaseDatos", "Usuario")
End Function

Public Function ClaveBD() As String
    ClaveBD = LeerIni("BaseDatos", "Clave")
End Function

Public Function ConnectionTimeout() As Long
    ConnectionTimeout = CLng(Val(LeerIni("BaseDatos", "ConnectionTimeout", "15")))
    If ConnectionTimeout <= 0 Then ConnectionTimeout = 15
End Function

Public Function CommandTimeout() As Long
    CommandTimeout = CLng(Val(LeerIni("BaseDatos", "CommandTimeout", "60")))
    If CommandTimeout <= 0 Then CommandTimeout = 60
End Function

Public Function RutaReportes() As String
    RutaReportes = ResolverRuta(LeerIni("Reportes", "RutaReportes", "."))
End Function

Public Function RutaPDF() As String
    RutaPDF = ResolverRuta(LeerIni("Reportes", "RutaPDF", "."))
End Function

Public Function RutaLog() As String
    RutaLog = ResolverRuta(LeerIni("Log", "RutaLog", "."))
End Function

' Ruta absoluta (C:\..., \\servidor\...) o relativa a App.Path
Private Function ResolverRuta(ByVal Ruta As String) As String
    If Len(Ruta) = 0 Then Ruta = "."
    If Mid$(Ruta, 2, 1) = ":" Or Left$(Ruta, 2) = "\\" Then
        ResolverRuta = Ruta
    Else
        ResolverRuta = App.Path & "\" & Ruta
    End If
    If Right$(ResolverRuta, 1) = "\" Then ResolverRuta = Left$(ResolverRuta, Len(ResolverRuta) - 1)
End Function
