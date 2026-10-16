VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10950
   LinkTopic       =   "Form2"
   ScaleHeight     =   8055
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   7095
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   9255
      _cx             =   16325
      _cy             =   12515
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1033
      EnableInteractiveParameterPrompting=   0   'False
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private crApp As New CRAXDRT.Application
Private crRep As CRAXDRT.Report

Private Sub Form_Load()
    Dim t As CRAXDRT.DatabaseTable
    On Error GoTo ErrHandler

    Set crRep = crApp.OpenReport(App.Path & "\Clientes.rpt")

    For Each t In crRep.Database.Tables
        t.SetLogOnInfo ".\SQLEXPRESS", "Prueba", "sa", "Prueba2026!"
    Next

    CRViewer1.ReportSource = crRep
    CRViewer1.ViewReport
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        CRViewer1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    End If
End Sub
