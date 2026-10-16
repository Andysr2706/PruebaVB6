VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   13695
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   5520
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   8895
   End
   Begin VB.CommandButton VerReporte 
      Caption         =   "VerReporte"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton ProbarSQL 
      Caption         =   "ProbarSQL"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CN_STR As String = "Provider=SQLOLEDB;Data Source=.\SQLEXPRESS;" & _
    "Initial Catalog=Prueba;User ID=sa;Password=Prueba2026!;"

Private Sub ProbarSQL_Click()
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    On Error GoTo ErrHandler

    Set cn = New ADODB.Connection
    cn.Open CN_STR
    Set rs = cn.Execute("SELECT IdCliente, Nombre, NIT, Saldo FROM Clientes")

    List1.Clear
    Do While Not rs.EOF
        List1.AddItem rs!IdCliente & " | " & rs!Nombre & " | " & rs!NIT & " | " & _
                      Format(rs!Saldo, "#,##0.00")
        rs.MoveNext
    Loop

    rs.Close
    cn.Close
    MsgBox "Conexion a SQL Server OK", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Private Sub VerReporte_Click()
    Form2.Show
End Sub
