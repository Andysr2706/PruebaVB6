Attribute VB_Name = "modUtil"
'=====================================================================
' modUtil - manejo explicito de Null, fechas y montos independientes
' de la configuracion regional del equipo.
'=====================================================================
Option Explicit

' Si recibe un ADODB.Field (rs!Campo) devuelve su .Value; si no, el valor.
Private Function ValorDe(ByVal x As Variant) As Variant
    If IsObject(x) Then ValorDe = x.Value Else ValorDe = x
End Function

Public Function TextoVar(ByVal x As Variant) As String
    x = ValorDe(x)
    If IsNull(x) Or IsEmpty(x) Then TextoVar = "Null" Else TextoVar = CStr(x)
End Function

Public Function NzCur(ByVal v As Variant, Optional ByVal Defecto As Currency = 0) As Currency
    v = ValorDe(v)
    If IsNull(v) Or IsEmpty(v) Then NzCur = Defecto Else NzCur = CCur(v)
End Function

Public Function NzLng(ByVal v As Variant, Optional ByVal Defecto As Long = 0) As Long
    v = ValorDe(v)
    If IsNull(v) Or IsEmpty(v) Then NzLng = Defecto Else NzLng = CLng(v)
End Function

Public Function NzStr(ByVal v As Variant, Optional ByVal Defecto As String = "") As String
    v = ValorDe(v)
    If IsNull(v) Or IsEmpty(v) Then NzStr = Defecto Else NzStr = CStr(v)
End Function

' dd/mm/aaaa literal (la "/" escapada no depende del separador regional)
Public Function FechaTexto(ByVal v As Variant) As String
    v = ValorDe(v)
    If IsNull(v) Or IsEmpty(v) Then
        FechaTexto = ""
    Else
        FechaTexto = Format$(CDate(v), "dd\/mm\/yyyy")
    End If
End Function

Public Function MontoTexto(ByVal v As Variant) As String
    v = ValorDe(v)
    If IsNull(v) Or IsEmpty(v) Then
        MontoTexto = ""
    Else
        MontoTexto = Format$(CCur(v), "#,##0.00")
    End If
End Function

Private Function SoloDigitos(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next
    SoloDigitos = True
End Function

' Acepta exactamente dd/mm/aaaa y valida que la fecha exista.
Public Function ParseFecha(ByVal Texto As String, ByRef Valor As Date) As Boolean
    Dim p() As String
    Dim d As Integer, m As Integer, a As Integer
    p = Split(Trim$(Texto), "/")
    If UBound(p) <> 2 Then Exit Function
    If Not SoloDigitos(p(0)) Or Not SoloDigitos(p(1)) Or Not SoloDigitos(p(2)) Then Exit Function
    If Len(p(2)) <> 4 Or Len(p(0)) > 2 Or Len(p(1)) > 2 Then Exit Function
    d = CInt(p(0)): m = CInt(p(1)): a = CInt(p(2))
    If m < 1 Or m > 12 Or d < 1 Or a < 1900 Or a > 2100 Then Exit Function
    If Day(DateSerial(a, m, d)) <> d Then Exit Function   ' ej. 31/02
    Valor = DateSerial(a, m, d)
    ParseFecha = True
End Function

' Monto positivo con punto decimal y maximo 2 decimales; acepta comas de miles.
' No usa CDbl/Val: evita punto flotante y dependencia regional.
Public Function ParseMonto(ByVal Texto As String, ByRef Valor As Currency) As Boolean
    Dim sEnt As String, sDec As String
    Dim p As Long
    Texto = Replace(Trim$(Texto), ",", "")
    If Len(Texto) = 0 Then Exit Function
    p = InStr(Texto, ".")
    If p > 0 Then
        sEnt = Left$(Texto, p - 1)
        sDec = Mid$(Texto, p + 1)
    Else
        sEnt = Texto
    End If
    If Len(sEnt) = 0 Then sEnt = "0"
    If Not SoloDigitos(sEnt) Or Len(sEnt) > 13 Then Exit Function
    If Len(sDec) > 0 Then
        If Not SoloDigitos(sDec) Or Len(sDec) > 2 Then Exit Function
    End If
    sDec = Left$(sDec & "00", 2)
    Valor = CCur(sEnt) + CCur(sDec) / 100
    ParseMonto = True
End Function
