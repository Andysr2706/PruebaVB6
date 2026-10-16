Attribute VB_Name = "modReportes"
'=====================================================================
' modReportes - integracion con Crystal Reports XI R2 (RDC craxdrt.dll)
' - Reasigna la conexion en tiempo de ejecucion desde config.ini
'   (reporte principal y subreportes): los .rpt no dependen del
'   servidor/base usados en diseno.
' - Envia parametros desde codigo sin dialogo de Crystal.
' - Exporta a PDF sin intervencion del usuario.
'=====================================================================
Option Explicit

Private Const MODULO As String = "modReportes"
Private mApp As CRAXDRT.Application

Public Const RPT_ESTADO_CUENTA As String = "EstadoCuenta.rpt"
Public Const RPT_ANTIGUEDAD As String = "AntiguedadSaldos.rpt"

Private Function CrApp() As CRAXDRT.Application
    If mApp Is Nothing Then Set mApp = New CRAXDRT.Application
    Set CrApp = mApp
End Function

Public Sub LiberarCrystal()
    Set mApp = Nothing
End Sub

Public Function AbrirReporte(ByVal Archivo As String) As CRAXDRT.Report
    On Error GoTo EH
    Dim sRuta As String
    Dim rep As CRAXDRT.Report
    sRuta = modConfig.RutaReportes() & "\" & Archivo
    If Len(Dir$(sRuta)) = 0 Then
        Err.Raise vbObjectError + 3000, MODULO, "No se encontro el reporte: " & sRuta
    End If
    Set rep = CrApp().OpenReport(sRuta, 1)    ' 1 = crOpenReportByTempCopy
    rep.EnableParameterPrompting = False
    rep.DiscardSavedData
    Set AbrirReporte = rep
    Set rep = Nothing
    Exit Function
EH:
    Set rep = Nothing
    Relanzar "AbrirReporte"
End Function

' Reasigna servidor/base/credenciales de config.ini en el reporte y
' en todos sus subreportes.
Public Sub AsignarConexion(ByVal rep As CRAXDRT.Report)
    On Error GoTo EH
    Dim sec As CRAXDRT.Section
    Dim obj As Object
    Dim subRep As CRAXDRT.Report

    AsignarLogOn rep
    For Each sec In rep.Sections
        For Each obj In sec.ReportObjects
            If obj.Kind = crSubreportObject Then
                Set subRep = rep.OpenSubreport(obj.SubreportName)
                AsignarLogOn subRep
                Set subRep = Nothing
            End If
        Next
    Next
    Exit Sub
EH:
    Set subRep = Nothing
    Relanzar "AsignarConexion"
End Sub

Private Sub AsignarLogOn(ByVal rep As CRAXDRT.Report)
    Dim t As CRAXDRT.DatabaseTable
    For Each t In rep.Database.Tables
        t.SetLogOnInfo modConfig.Servidor(), modConfig.BaseDatos(), modConfig.UsuarioBD(), modConfig.ClaveBD()
    Next
End Sub

' Busca el parametro por nombre ("@ClienteId" o "{?@ClienteId}") y asigna su valor.
Public Sub AsignarParametro(ByVal rep As CRAXDRT.Report, ByVal Nombre As String, ByVal Valor As Variant)
    On Error GoTo EH
    Dim p As CRAXDRT.ParameterFieldDefinition
    Dim bEncontrado As Boolean
    For Each p In rep.ParameterFields
        If StrComp(p.ParameterFieldName, Nombre, vbTextCompare) = 0 Or _
           StrComp(p.Name, "{?" & Nombre & "}", vbTextCompare) = 0 Then
            p.ClearCurrentValueAndRange
            p.AddCurrentValue Valor
            bEncontrado = True
        End If
    Next
    If Not bEncontrado Then
        Err.Raise vbObjectError + 3001, MODULO, "El reporte no tiene el parametro " & Nombre
    End If
    Exit Sub
EH:
    Relanzar "AsignarParametro"
End Sub

' Exporta a PDF en la ruta de config.ini, sin dialogos. Devuelve la ruta.
Public Function ExportarPDF(ByVal rep As CRAXDRT.Report, ByVal NombreBase As String) As String
    On Error GoTo EH
    Dim sCarpeta As String, sArchivo As String
    sCarpeta = modConfig.RutaPDF()
    If Len(Dir$(sCarpeta, vbDirectory)) = 0 Then MkDir sCarpeta
    sArchivo = sCarpeta & "\" & NombreBase & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"
    With rep.ExportOptions
        .DestinationType = crEDTDiskFile
        .FormatType = crEFTPortableDocFormat
        .DiskFileName = sArchivo
        .PDFExportAllPages = True      ' sin esto Crystal espera un rango de paginas (dialogo oculto)
    End With
    rep.Export False
    ExportarPDF = sArchivo
    Exit Function
EH:
    Relanzar "ExportarPDF"
End Function

Private Sub Relanzar(ByVal Procedimiento As String)
    Dim lNum As Long, sDesc As String, sFuente As String
    lNum = Err.Number: sDesc = Err.Description: sFuente = Err.Source
    If sFuente <> modLog.FUENTE_REGISTRADO Then modLog.RegistrarError MODULO, Procedimiento, lNum, sDesc
    Err.Raise lNum, modLog.FUENTE_REGISTRADO, sDesc
End Sub
