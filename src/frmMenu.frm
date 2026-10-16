VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobranza Express"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "frmMenu"
   MaxButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   3360
      Width           =   3255
   End
   Begin VB.CommandButton cmdPruebas 
      Caption         =   "Pruebas de interes &moratorio"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   2640
      Width           =   3255
   End
   Begin VB.CommandButton cmdReportes 
      Caption         =   "&Reportes"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton cmdPago 
      Caption         =   "Registro de &pago"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton cmdClientes 
      Caption         =   "Consulta de &clientes con saldo"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULO As String = "frmMenu"

Private Sub Form_Load()
    On Error GoTo EH
    If Not modConfig.ConfigExiste() Then
        MsgBox "No se encontro config.ini en " & App.Path & vbCrLf & _
               "Copie config.ini.example como config.ini y complete la clave.", vbCritical, "Cobranza Express"
    End If
    Exit Sub
EH:
    modLog.MostrarError MODULO, "Form_Load", Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdClientes_Click()
    frmClientes.Show
End Sub

Private Sub cmdPago_Click()
    frmPago.Show
End Sub

Private Sub cmdReportes_Click()
    frmReportes.Show
End Sub

Private Sub cmdPruebas_Click()
    frmPruebasMora.Show
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim f As Form
    For Each f In Forms
        If Not f Is Me Then Unload f
    Next
    modReportes.LiberarCrystal
End Sub
