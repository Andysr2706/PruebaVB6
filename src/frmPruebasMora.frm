VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPruebasMora 
   Caption         =   "Pruebas de interes moratorio (VB6 vs SQL Server)"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14400
   LinkTopic       =   "frmPruebasMora"
   ScaleHeight     =   6600
   ScaleWidth      =   14400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEjecutar 
      Caption         =   "&Ejecutar pruebas"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid grdPruebas 
      Height          =   5295
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   9340
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Label lblResumen 
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   300
      Width           =   11415
   End
End
Attribute VB_Name = "frmPruebasMora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Rutina de pruebas: cada caso muestra el esperado, el obtenido en VB6
' (clsCalculoMora) y el obtenido en SQL Server (tvf_Mora via SP).
' Fecha de corte fija para que los resultados sean reproducibles.
Option Explicit

Private Const MODULO As String = "frmPruebasMora"
Private mCorte As Date
Private mOK As Long, mTotal As Long

Private Sub Form_Load()
    mCorte = DateSerial(2026, 10, 16)
    modGrid.PrepararGrid grdPruebas, "#|Caso|Saldo|Tasa|Gracia|Vencimiento|Esperado|VB6|SQL|Senal VB6|Senal SQL|Resultado", _
                         "400|3600|1100|800|700|1200|1000|1000|1000|1500|1500|1000"
End Sub

Private Sub cmdEjecutar_Click()
    On Error GoTo EH
    Dim neg As clsCobranza
    Dim mora As clsCalculoMora
    Dim vDef As Variant

    Screen.MousePointer = vbHourglass
    Set neg = New clsCobranza
    Set mora = New clsCalculoMora
    mOK = 0: mTotal = 0
    Form_Load
    vDef = mora.TasaEfectiva(Null, (CDec(18) / 100))   ' tasa por defecto de prueba

    '      descripcion                                         saldo     tasa              gracia venc                      esperado
    Caso neg, mora, "Vencida 30 dias, 24%, sin gracia", CCur(1000), (CDec(24) / 100), 0, mCorte - 30, CCur(19.73)
    Caso neg, mora, "Con 5 dias de gracia (25 dias mora)", CCur(1000), (CDec(24) / 100), 5, mCorte - 30, CCur(16.44)
    Caso neg, mora, "Gracia mayor que el atraso => 0", CCur(1000), (CDec(24) / 100), 40, mCorte - 30, CCur(0)
    Caso neg, mora, "Dias de gracia Null => 0 dias", CCur(1000), (CDec(24) / 100), Null, mCorte - 30, CCur(19.73)
    Caso neg, mora, "Vencimiento Null => sin interes, marcada", CCur(1000), (CDec(24) / 100), 0, Null, CCur(0)
    Caso neg, mora, "Tasa cliente y defecto Null => sin interes", CCur(1000), mora.TasaEfectiva(Null, Null), 0, mCorte - 30, CCur(0)
    Caso neg, mora, "Tasa cliente Null => usa defecto 18%", CCur(1000), vDef, 0, mCorte - 30, CCur(14.79)
    Caso neg, mora, "Nota de credito (saldo negativo) => 0", CCur(-500), (CDec(24) / 100), 0, mCorte - 30, CCur(0)
    Caso neg, mora, "No vencida (vence en 10 dias) => 0", CCur(1000), (CDec(24) / 100), 0, mCorte + 10, CCur(0)
    Caso neg, mora, "Redondeo comercial 2.125 => 2.13", CCur(3102.5), (CDec(25) / 100), 0, mCorte - 1, CCur(2.13)
    Caso neg, mora, "Redondeo comercial 2.135 => 2.14", CCur(3117.1), (CDec(25) / 100), 0, mCorte - 1, CCur(2.14)
    Caso neg, mora, "Redondeo 0.005 => 0.01", CCur(7.3), (CDec(25) / 100), 0, mCorte - 1, CCur(0.01)

    lblResumen.Caption = mOK & " de " & mTotal & " casos correctos (esperado = VB6 = SQL). " & _
                         "Referencia: Round(2.125, 2) de VB6 da " & Round(2.125, 2) & " (bancario), por eso no se usa."
Salir:
    Screen.MousePointer = vbDefault
    Set mora = Nothing
    Set neg = Nothing
    Exit Sub
EH:
    Screen.MousePointer = vbDefault
    modLog.MostrarError MODULO, "cmdEjecutar_Click", Err.Number, Err.Description, Err.Source
    Resume Salir
End Sub

Private Sub Caso(ByVal neg As clsCobranza, ByVal mora As clsCalculoMora, ByVal Descripcion As String, _
                 ByVal Saldo As Currency, ByVal Tasa As Variant, ByVal Gracia As Variant, _
                 ByVal Venc As Variant, ByVal Esperado As Currency)
    Dim cVB As Currency, cSQL As Currency
    Dim lDiasSQL As Long, sSenalSQL As String
    Dim bOK As Boolean

    cVB = mora.Calcular(Saldo, Tasa, Gracia, Venc, mCorte)
    cSQL = neg.MoraSQL(Saldo, Tasa, Gracia, Venc, mCorte, lDiasSQL, sSenalSQL)
    bOK = (cVB = Esperado And cSQL = Esperado And mora.Senal = sSenalSQL)
    mTotal = mTotal + 1
    If bOK Then mOK = mOK + 1

    modGrid.AgregarFila grdPruebas, mTotal, Descripcion, modUtil.MontoTexto(Saldo), _
        modUtil.TextoVar(Tasa), modUtil.TextoVar(Gracia), _
        IIf(IsNull(Venc), "Null", modUtil.FechaTexto(Venc)), modUtil.MontoTexto(Esperado), _
        modUtil.MontoTexto(cVB), modUtil.MontoTexto(cSQL), mora.Senal, sSenalSQL, IIf(bOK, "OK", "FALLA")
End Sub
