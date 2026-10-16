Attribute VB_Name = "modGrid"
' modGrid - utilidades de presentacion para MSFlexGrid
Option Explicit

' Encabezados separados por "|" y anchos en twips separados por "|"
Public Sub PrepararGrid(ByVal g As Object, ByVal Encabezados As String, ByVal Anchos As String)
    Dim aE() As String, aA() As String
    Dim i As Long
    aE = Split(Encabezados, "|")
    aA = Split(Anchos, "|")
    g.Clear
    g.FixedCols = 0
    g.Rows = 2
    g.FixedRows = 1
    g.Cols = UBound(aE) + 1
    For i = 0 To UBound(aE)
        g.TextMatrix(0, i) = aE(i)
        If i <= UBound(aA) Then g.ColWidth(i) = CLng(aA(i))
    Next
    g.Rows = 1
End Sub

Public Sub AgregarFila(ByVal g As Object, ParamArray Valores() As Variant)
    Dim i As Long, r As Long
    g.Rows = g.Rows + 1
    r = g.Rows - 1
    For i = 0 To UBound(Valores)
        If i < g.Cols Then g.TextMatrix(r, i) = CStr(Valores(i))
    Next
End Sub
