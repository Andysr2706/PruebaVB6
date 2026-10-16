VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmClientes 
   Caption         =   "Consulta de clientes con saldo"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14400
   LinkTopic       =   "frmClientes"
   ScaleHeight     =   8400
   ScaleWidth      =   14400
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBuscar 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin VB.TextBox txtFechaCorte 
      Height          =   315
      Left            =   8640
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Default         =   -1  'True
      Height          =   375
      Left            =   10200
      TabIndex        =   2
      Top             =   210
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid grdClientes 
      Height          =   3735
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   6588
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid grdFacturas 
      Height          =   3135
      Left            =   240
      TabIndex        =   4
      Top             =   4920
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   5530
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label lblFacturas 
      Caption         =   "Facturas pendientes (seleccione un cliente)"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4560
      Width           =   13935
   End
   Begin VB.Label lblFecha 
      Caption         =   "Fecha corte (dd/mm/aaaa):"
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   285
      Width           =   1815
   End
   Begin VB.Label lblBuscar 
      Caption         =   "Codigo, nombre o NIT:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   285
      Width           =   2055
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULO As String = "frmClientes"
Private mFechaCorte As Date

Private Sub Form_Load()
    txtFechaCorte.Text = modUtil.FechaTexto(Date)
    modGrid.PrepararGrid grdClientes, "Id|Codigo|Nombre|NIT|Limite credito|Saldo total|Saldo vencido|Dias atraso|Interes moratorio|Saldo a favor|Observacion", _
                         "0|900|3200|1200|1300|1300|1300|1000|1400|1200|1400"
    modGrid.PrepararGrid grdFacturas, "Numero|Tipo|Emision|Vencimiento|Monto|Saldo|Dias mora|Interes (VB6)|Interes (SQL)|Coincide|Senal", _
                         "1100|500|1100|1200|1300|1300|900|1300|1300|900|1700"
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo EH
    Dim neg As clsCobranza
    Dim rs As ADODB.Recordset

    If Not modUtil.ParseFecha(txtFechaCorte.Text, mFechaCorte) Then
        MsgBox "Fecha de corte invalida. Use el formato dd/mm/aaaa.", vbExclamation
        txtFechaCorte.SetFocus
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    Set neg = New clsCobranza
    Set rs = neg.BuscarClientes(txtBuscar.Text, mFechaCorte)

    modGrid.PrepararGrid grdClientes, "Id|Codigo|Nombre|NIT|Limite credito|Saldo total|Saldo vencido|Dias atraso|Interes moratorio|Saldo a favor|Observacion", _
                         "0|900|3200|1200|1300|1300|1300|1000|1400|1200|1400"
    modGrid.PrepararGrid grdFacturas, "Numero|Tipo|Emision|Vencimiento|Monto|Saldo|Dias mora|Interes (VB6)|Interes (SQL)|Coincide|Senal", _
                         "1100|500|1100|1200|1300|1300|900|1300|1300|900|1700"
    Do While Not rs.EOF
        modGrid.AgregarFila grdClientes, modUtil.NzLng(rs!ClienteId), modUtil.NzStr(rs!Codigo), _
            modUtil.NzStr(rs!Nombre), modUtil.NzStr(rs!NIT, "(sin NIT)"), _
            modUtil.MontoTexto(rs!LimiteCredito), modUtil.MontoTexto(rs!SaldoTotal), _
            modUtil.MontoTexto(rs!SaldoVencido), modUtil.NzLng(rs!DiasMayorAtraso), _
            modUtil.MontoTexto(rs!InteresMoratorio), modUtil.MontoTexto(rs!SaldoAFavor), _
            modUtil.NzStr(rs!Observacion)
        rs.MoveNext
    Loop
    lblFacturas.Caption = rs.RecordCount & " cliente(s). Seleccione uno para ver sus facturas pendientes."

Salir:
    Screen.MousePointer = vbDefault
    If Not rs Is Nothing Then
        If rs.State <> adStateClosed Then rs.Close
    End If
    Set rs = Nothing
    Set neg = Nothing
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    modLog.MostrarError MODULO, "cmdBuscar_Click", Err.Number, Err.Description, Err.Source
    Resume Salir
End Sub

Private Sub grdClientes_Click()
    On Error GoTo EH
    Dim neg As clsCobranza
    Dim mora As clsCalculoMora
    Dim rs As ADODB.Recordset
    Dim lId As Long
    Dim cVB As Currency, cSQL As Currency

    If grdClientes.Row < 1 Then Exit Sub
    lId = CLng(Val(grdClientes.TextMatrix(grdClientes.Row, 0)))
    If lId = 0 Then Exit Sub

    Screen.MousePointer = vbHourglass
    Set neg = New clsCobranza
    Set mora = New clsCalculoMora
    Set rs = neg.FacturasPendientes(lId, mFechaCorte)

    modGrid.PrepararGrid grdFacturas, "Numero|Tipo|Emision|Vencimiento|Monto|Saldo|Dias mora|Interes (VB6)|Interes (SQL)|Coincide|Senal", _
                         "1100|500|1100|1200|1300|1300|900|1300|1300|900|1700"
    Do While Not rs.EOF
        cVB = mora.Calcular(modUtil.NzCur(rs!Saldo), rs!TasaAplicada.Value, rs!DiasGracia.Value, _
                            rs!FechaVencimiento.Value, mFechaCorte)
        cSQL = modUtil.NzCur(rs!Interes)
        modGrid.AgregarFila grdFacturas, modUtil.NzStr(rs!NumeroFactura), modUtil.NzStr(rs!Tipo), _
            modUtil.FechaTexto(rs!FechaEmision), IIf(IsNull(rs!FechaVencimiento.Value), "(sin venc.)", modUtil.FechaTexto(rs!FechaVencimiento)), _
            modUtil.MontoTexto(rs!Monto), modUtil.MontoTexto(rs!Saldo), mora.DiasMora, _
            modUtil.MontoTexto(cVB), modUtil.MontoTexto(cSQL), IIf(cVB = cSQL, "SI", "NO"), mora.Senal
        rs.MoveNext
    Loop
    lblFacturas.Caption = "Facturas pendientes de " & grdClientes.TextMatrix(grdClientes.Row, 2) & _
                          " (orden FIFO: vencimiento y numero)"

Salir:
    Screen.MousePointer = vbDefault
    If Not rs Is Nothing Then
        If rs.State <> adStateClosed Then rs.Close
    End If
    Set rs = Nothing
    Set mora = Nothing
    Set neg = Nothing
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    modLog.MostrarError MODULO, "grdClientes_Click", Err.Number, Err.Description, Err.Source
    Resume Salir
End Sub
