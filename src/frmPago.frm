VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de pago"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10800
   LinkTopic       =   "frmPago"
   MaxButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboCliente 
      Height          =   315
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin VB.TextBox txtFiltro 
      Height          =   315
      Left            =   1800
      TabIndex        =   15
      ToolTipText     =   "Escriba parte del codigo, nombre o NIT"
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtFecha 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ComboBox cboForma 
      Height          =   315
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtReferencia 
      Height          =   315
      Left            =   5520
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdVistaPrevia 
      Caption         =   "&Vista previa"
      Height          =   495
      Left            =   8160
      TabIndex        =   5
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar pago"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8160
      TabIndex        =   6
      Top             =   840
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid grdPrevia 
      Height          =   4335
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7646
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin VB.Label lblResultado 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   240
      TabIndex        =   13
      Top             =   6360
      Width           =   10215
   End
   Begin VB.Label lblPrevia 
      Caption         =   "Vista previa de aplicacion (FIFO)"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1620
      Width           =   10215
   End
   Begin VB.Label lblReferencia 
      Caption         =   "Referencia:"
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   1245
      Width           =   1455
   End
   Begin VB.Label lblForma 
      Caption         =   "Forma de pago:"
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   765
      Width           =   1455
   End
   Begin VB.Label lblMonto 
      Caption         =   "Monto:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1245
      Width           =   1455
   End
   Begin VB.Label lblFecha 
      Caption         =   "Fecha (dd/mm/aaaa):"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   765
      Width           =   1575
   End
   Begin VB.Label lblCliente 
      Caption         =   "Cliente (filtro):"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   285
      Width           =   1455
   End
End
Attribute VB_Name = "frmPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULO As String = "frmPago"
Private mClaveVistaPrevia As String   ' datos con los que se genero la vista previa
Private mClientes As Collection

Private Sub Form_Load()
    On Error GoTo EH
    Dim neg As clsCobranza
    Dim col As Collection
    Dim v As Variant

    txtFecha.Text = modUtil.FechaTexto(Date)
    cboForma.AddItem "EFECTIVO"
    cboForma.AddItem "CHEQUE"
    cboForma.AddItem "TRANSFERENCIA"
    cboForma.AddItem "TARJETA"
    cboForma.AddItem "DEPOSITO"
    PrepararGrid

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
    cmdAplicar.Enabled = False
End Sub

Private Sub PrepararGrid()
    modGrid.PrepararGrid grdPrevia, "Orden|Factura|Vencimiento|Saldo anterior|Monto a aplicar|Saldo nuevo", _
                         "700|1600|1500|2000|2000|2000"
End Sub

' Valida la captura y devuelve los valores ya convertidos.
Private Function LeerCaptura(ByRef ClienteId As Long, ByRef Fecha As Date, ByRef Monto As Currency) As Boolean
    If cboCliente.ListIndex < 0 Then
        MsgBox "Seleccione un cliente.", vbExclamation: cboCliente.SetFocus: Exit Function
    End If
    If Not modUtil.ParseFecha(txtFecha.Text, Fecha) Then
        MsgBox "Fecha invalida. Use dd/mm/aaaa.", vbExclamation: txtFecha.SetFocus: Exit Function
    End If
    If Fecha > Date Then
        MsgBox "La fecha del pago no puede ser futura.", vbExclamation: txtFecha.SetFocus: Exit Function
    End If
    If Not modUtil.ParseMonto(txtMonto.Text, Monto) Then
        MsgBox "Monto invalido. Use numeros con punto decimal (maximo 2 decimales).", vbExclamation
        txtMonto.SetFocus: Exit Function
    End If
    If Monto <= 0 Then
        MsgBox "El monto debe ser mayor que cero.", vbExclamation: txtMonto.SetFocus: Exit Function
    End If
    ClienteId = cboCliente.ItemData(cboCliente.ListIndex)
    LeerCaptura = True
End Function

Private Function ClaveActual() As String
    Dim lId As Long
    If cboCliente.ListIndex >= 0 Then lId = cboCliente.ItemData(cboCliente.ListIndex)
    ClaveActual = lId & "|" & txtFecha.Text & "|" & txtMonto.Text
End Function

Private Sub cmdVistaPrevia_Click()
    On Error GoTo EH
    Dim neg As clsCobranza
    Dim rs As ADODB.Recordset
    Dim lCli As Long, dFecha As Date, cMonto As Currency, cFavor As Currency

    cmdAplicar.Enabled = False
    If Not LeerCaptura(lCli, dFecha, cMonto) Then Exit Sub

    Screen.MousePointer = vbHourglass
    Set neg = New clsCobranza
    Set rs = neg.SimularPago(lCli, cMonto, cFavor)
    PrepararGrid
    Do While Not rs.EOF
        modGrid.AgregarFila grdPrevia, modUtil.NzLng(rs!Orden), modUtil.NzStr(rs!NumeroFactura), _
            IIf(IsNull(rs!FechaVencimiento.Value), "(sin venc.)", modUtil.FechaTexto(rs!FechaVencimiento)), _
            modUtil.MontoTexto(rs!SaldoAnterior), modUtil.MontoTexto(rs!MontoAplicado), modUtil.MontoTexto(rs!SaldoNuevo)
        rs.MoveNext
    Loop
    lblResultado.Caption = "Monto: " & modUtil.MontoTexto(cMonto) & _
                           "   Aplicado a facturas: " & modUtil.MontoTexto(cMonto - cFavor) & _
                           "   Saldo a favor resultante: " & modUtil.MontoTexto(cFavor)
    mClaveVistaPrevia = ClaveActual()
    cmdAplicar.Enabled = True

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
    modLog.MostrarError MODULO, "cmdVistaPrevia_Click", Err.Number, Err.Description, Err.Source
    Resume Salir
End Sub

Private Sub cmdAplicar_Click()
    On Error GoTo EH
    Dim neg As clsCobranza
    Dim lCli As Long, dFecha As Date, cMonto As Currency
    Dim lPago As Long, cFavor As Currency
    Dim eRes As ResultadoPago

    cmdAplicar.Enabled = False            ' evita doble clic
    If ClaveActual() <> mClaveVistaPrevia Then
        MsgBox "Los datos cambiaron. Genere de nuevo la vista previa.", vbExclamation
        Exit Sub
    End If
    If Not LeerCaptura(lCli, dFecha, cMonto) Then Exit Sub
    If MsgBox("Confirma aplicar el pago de " & modUtil.MontoTexto(cMonto) & " a " & cboCliente.Text & "?", _
              vbQuestion + vbYesNo, "Confirmar pago") = vbNo Then
        cmdAplicar.Enabled = True
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    Set neg = New clsCobranza
    eRes = neg.AplicarPago(lCli, dFecha, cMonto, cboForma.Text, txtReferencia.Text, lPago, cFavor)
    Screen.MousePointer = vbDefault

    If eRes = rpOK Then
        lblResultado.Caption = neg.MensajeResultado(eRes) & "  Pago No. " & lPago & _
                               "   Saldo a favor generado: " & modUtil.MontoTexto(cFavor)
        MsgBox lblResultado.Caption, vbInformation, "Cobranza Express"
        txtMonto.Text = ""
        txtReferencia.Text = ""
        mClaveVistaPrevia = ""
        PrepararGrid
    Else
        lblResultado.Caption = neg.MensajeResultado(eRes)
        MsgBox lblResultado.Caption, vbExclamation, "Cobranza Express"
    End If

Salir:
    Screen.MousePointer = vbDefault
    Set neg = Nothing
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    modLog.MostrarError MODULO, "cmdAplicar_Click", Err.Number, Err.Description, Err.Source
    Resume Salir
End Sub

' Cualquier cambio invalida la vista previa
Private Sub cboCliente_Click()
    cmdAplicar.Enabled = False
End Sub

Private Sub txtFecha_Change()
    cmdAplicar.Enabled = False
End Sub

Private Sub txtMonto_Change()
    cmdAplicar.Enabled = False
End Sub
