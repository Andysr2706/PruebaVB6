VERSION 5.00
Begin VB.Form frmReportes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8400
   LinkTopic       =   "frmReportes"
   MaxButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optEstadoCuenta 
      Caption         =   "Estado de cuenta por cliente (modelo pull: SP)"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   5295
   End
   Begin VB.OptionButton optAntiguedad 
      Caption         =   "Antiguedad de saldos (modelo push: Recordset ADO)"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   5295
   End
   Begin VB.ComboBox cboCliente 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox txtFiltro 
      Height          =   315
      Left            =   2280
      TabIndex        =   9
      ToolTipText     =   "Escriba parte del codigo, nombre o NIT"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtFechaCorte 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "&Ver en pantalla"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton cmdPDF 
      Caption         =   "&Exportar a PDF"
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label lblEstado 
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2940
      Width           =   7935
   End
   Begin VB.Label lblFecha 
      Caption         =   "Fecha de corte:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1725
      Width           =   1935
   End
   Begin VB.Label lblCliente 
      Caption         =   "Cliente (filtro):"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1245
      Width           =   1935
   End
End
Attribute VB_Name = "frmReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULO As String = "frmReportes"
Private mClientes As Collection

Private Sub Form_Load()
    On Error GoTo EH
    Dim neg As clsCobranza
    Dim col As Collection
    Dim v As Variant
    txtFechaCorte.Text = modUtil.FechaTexto(Date)
    Set neg = New clsCobranza
    Set mClientes = neg.ListaClientes()
    LlenarCombo ""
Salir:
    Set col = Nothing
    Set neg = Nothing
    Exit Sub
EH:
    modLog.MostrarError MODULO, "Form_Load", Err.Number, Err.Description, Err.Source
    Resume Salir
End Sub

' Filtro local sobre la lista cargada una sola vez (cursor forward-only)
Private Sub LlenarCombo(ByVal Filtro As String)
    Dim v As Variant
    cboCliente.Clear
    If mClientes Is Nothing Then Exit Sub
    For Each v In mClientes
        If Len(Filtro) = 0 Or InStr(1, v(1), Filtro, vbTextCompare) > 0 Then
            cboCliente.AddItem v(1)
            cboCliente.ItemData(cboCliente.NewIndex) = v(0)
        End If
    Next
    If cboCliente.ListCount = 1 Then cboCliente.ListIndex = 0
End Sub

Private Sub txtFiltro_Change()
    LlenarCombo Trim$(txtFiltro.Text)
End Sub

Private Sub optEstadoCuenta_Click()
    cboCliente.Enabled = True
    txtFiltro.Enabled = True
End Sub

Private Sub optAntiguedad_Click()
    cboCliente.Enabled = False
    txtFiltro.Enabled = False
End Sub

' Prepara el reporte elegido con conexion y parametros desde codigo.
Private Function PrepararReporte(ByRef NombreBase As String) As CRAXDRT.Report
    Dim dCorte As Date
    Dim rep As CRAXDRT.Report
    Dim neg As clsCobranza
    Dim rs As ADODB.Recordset

    If Not modUtil.ParseFecha(txtFechaCorte.Text, dCorte) Then
        Err.Raise vbObjectError + 4000, modLog.FUENTE_REGISTRADO, "Fecha de corte invalida. Use dd/mm/aaaa."
    End If

    If optEstadoCuenta.Value Then
        If cboCliente.ListIndex < 0 Then
            Err.Raise vbObjectError + 4001, modLog.FUENTE_REGISTRADO, "Seleccione un cliente."
        End If
        ' PULL: Crystal ejecuta usp_Reporte_EstadoCuenta con la conexion de config.ini
        Set rep = modReportes.AbrirReporte(modReportes.RPT_ESTADO_CUENTA)
        modReportes.AsignarConexion rep
        modReportes.AsignarParametro rep, "@ClienteId", CLng(cboCliente.ItemData(cboCliente.ListIndex))
        modReportes.AsignarParametro rep, "@FechaCorte", dCorte
        NombreBase = "EstadoCuenta_" & cboCliente.ItemData(cboCliente.ListIndex)
    Else
        ' PUSH: VB6 obtiene el Recordset (capa de negocio/datos) y lo entrega al reporte
        Set neg = New clsCobranza
        Set rs = neg.DatosAntiguedad(dCorte)
        Set rep = modReportes.AbrirReporte(modReportes.RPT_ANTIGUEDAD)
        ' El parametro se asigna ANTES de SetDataSource, que descarta los
        ' parametros del SP usado en diseno (por eso aqui es opcional).
        modReportes.AsignarParametro rep, "@FechaCorte", dCorte, False
        rep.Database.SetDataSource rs, 3, 1
        modReportes.AsignarFormulaTexto rep, "FechaCorteTexto", modUtil.FechaTexto(dCorte)
        NombreBase = "AntiguedadSaldos"
        Set rs = Nothing      ' el reporte conserva su propia referencia
        Set neg = Nothing
    End If
    Set PrepararReporte = rep
    Set rep = Nothing
End Function

Private Sub cmdVer_Click()
    On Error GoTo EH
    Dim rep As CRAXDRT.Report
    Dim sNombre As String
    Dim f As frmVisor
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Generando reporte..."
    Set rep = PrepararReporte(sNombre)
    Set f = New frmVisor
    f.Mostrar rep, sNombre
    lblEstado.Caption = "Reporte generado: " & sNombre
Salir:
    Screen.MousePointer = vbDefault
    Set rep = Nothing
    Set f = Nothing
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    lblEstado.Caption = ""
    modLog.MostrarError MODULO, "cmdVer_Click", Err.Number, Err.Description, Err.Source
    Resume Salir
End Sub

Private Sub cmdPDF_Click()
    On Error GoTo EH
    Dim rep As CRAXDRT.Report
    Dim sNombre As String, sArchivo As String
    Screen.MousePointer = vbHourglass
    lblEstado.Caption = "Exportando a PDF..."
    Set rep = PrepararReporte(sNombre)
    sArchivo = modReportes.ExportarPDF(rep, sNombre)
    lblEstado.Caption = "PDF generado: " & sArchivo
Salir:
    Screen.MousePointer = vbDefault
    Set rep = Nothing
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    lblEstado.Caption = ""
    modLog.MostrarError MODULO, "cmdPDF_Click", Err.Number, Err.Description, Err.Source
    Resume Salir
End Sub
