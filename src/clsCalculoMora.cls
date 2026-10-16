VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCalculoMora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================================
' clsCalculoMora - interes moratorio, independiente de BD e interfaz.
' Replica EXACTAMENTE dbo.tvf_Mora:
'   diasMora = max(0, fechaCorte - fechaVencimiento - diasGracia)
'   interes  = RedondeoComercial(saldo * tasa * diasMora / 365, 2)
' Se usa Decimal (Variant/CDec, 28 digitos) para el calculo intermedio
' y Currency para el resultado: nunca Double.
'=====================================================================
Option Explicit

Private mDiasMora As Long
Private mDiasAtraso As Long
Private mSenal As String

Public Property Get DiasMora() As Long
    DiasMora = mDiasMora
End Property

Public Property Get DiasAtraso() As Long
    DiasAtraso = mDiasAtraso
End Property

' "", "NOTA CREDITO", "SIN VENCIMIENTO", "SIN TASA" (mismos textos que SQL)
Public Property Get Senal() As String
    Senal = mSenal
End Property

' Tasa del cliente; si es Null usa la de defecto; si ambas faltan => Null
Public Function TasaEfectiva(ByVal TasaCliente As Variant, ByVal TasaDefecto As Variant) As Variant
    If Not IsNull(TasaCliente) And Not IsEmpty(TasaCliente) Then
        TasaEfectiva = CDec(TasaCliente)
    ElseIf Not IsNull(TasaDefecto) And Not IsEmpty(TasaDefecto) Then
        TasaEfectiva = CDec(TasaDefecto)
    Else
        TasaEfectiva = Null
    End If
End Function

Public Function Calcular(ByVal Saldo As Currency, ByVal Tasa As Variant, ByVal DiasGracia As Variant, _
                         ByVal FechaVencimiento As Variant, ByVal FechaCorte As Date) As Currency
    On Error GoTo EH
    Dim lGracia As Long
    Dim lDif As Long
    Dim vInteres As Variant

    mDiasMora = 0
    mDiasAtraso = 0
    mSenal = ""
    Calcular = 0

    If Saldo < 0 Then mSenal = "NOTA CREDITO": Exit Function
    If IsNull(FechaVencimiento) Or IsEmpty(FechaVencimiento) Then mSenal = "SIN VENCIMIENTO": Exit Function
    If IsNull(Tasa) Or IsEmpty(Tasa) Then mSenal = "SIN TASA"
    If Saldo = 0 Then Exit Function

    If IsNull(DiasGracia) Or IsEmpty(DiasGracia) Then lGracia = 0 Else lGracia = CLng(DiasGracia)
    lDif = DateDiff("d", CDate(FechaVencimiento), FechaCorte)
    If lDif > 0 Then mDiasAtraso = lDif
    If lDif - lGracia > 0 Then mDiasMora = lDif - lGracia

    If mDiasMora = 0 Or Len(mSenal) > 0 Then Exit Function

    ' Multiplicar primero (exacto) y dividir al final, igual que en SQL
    vInteres = CDec(Saldo) * CDec(Tasa) * CDec(mDiasMora)
    vInteres = vInteres / CDec(365)
    Calcular = CCur(RedondeoComercial(vInteres, 2))
    Exit Function
EH:
    Err.Raise Err.Number, "clsCalculoMora.Calcular", Err.Description
End Function

' Redondeo comercial: 0.5 se aleja de cero (2.125 -> 2.13, -2.125 -> -2.13).
' VB6 Round() usa redondeo bancario (Round(2.125, 2) = 2.12): no se usa.
Public Function RedondeoComercial(ByVal Valor As Variant, ByVal Decimales As Integer) As Variant
    Dim vFactor As Variant
    Dim vAbs As Variant
    vFactor = CDec(10 ^ Decimales)
    vAbs = CDec(Abs(CDec(Valor)))
    vAbs = Int(vAbs * vFactor + CDec(0.5)) / vFactor
    If CDec(Valor) < 0 Then RedondeoComercial = -vAbs Else RedondeoComercial = vAbs
End Function
