VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form frmVisor 
   Caption         =   "Visor de reportes"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12000
   LinkTopic       =   "frmVisor"
   ScaleHeight     =   8055
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
      EnableRefreshButton=   0   'False
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
      EnableLogonPrompts=   0   'False
      LocaleID        =   1033
      EnableInteractiveParameterPrompting=   0   'False
   End
End
Attribute VB_Name = "frmVisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULO As String = "frmVisor"
Private mRep As CRAXDRT.Report

Public Sub Mostrar(ByVal rep As CRAXDRT.Report, ByVal Titulo As String)
    On Error GoTo EH
    Set mRep = rep
    Me.Caption = Titulo
    CRViewer1.ReportSource = mRep
    CRViewer1.ViewReport
    Me.Show
    Exit Sub
EH:
    modLog.MostrarError MODULO, "Mostrar", Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then CRViewer1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mRep = Nothing
End Sub
