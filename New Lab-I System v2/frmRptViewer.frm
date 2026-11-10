VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Begin VB.Form frmRptViewer 
   BorderStyle     =   0  'None
   Caption         =   "Report Viewer"
   ClientHeight    =   3435
   ClientLeft      =   4365
   ClientTop       =   3210
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   -90
      Width           =   4785
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   4440
         TabIndex        =   3
         Top             =   180
         Width           =   255
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Viewer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   150
         Width           =   1410
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   420
      Width           =   4785
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
   End
End
Attribute VB_Name = "frmRptViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnDirectPrint As Boolean

Private Sub CRViewer1_CloseButtonClicked(UseDefault As Boolean)
    Unload Me
End Sub

Private Sub Form_Load()
CRViewer1.ReportSource = Report

If blnDirectPrint Then
   CRViewer1.PrintReport
Else
   CRViewer1.ViewReport
End If
End Sub

Private Sub Form_Resize()
CRViewer1.Top = fraTitle.Height

CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight - fraTitle.Height
CRViewer1.Width = ScaleWidth
fraTitle.Width = ScaleWidth
lblClose.Left = ScaleWidth - 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Report = Nothing
End Sub

Public Property Let directPrint(ByVal blnNewValue As Boolean)
   blnDirectPrint = blnNewValue
End Property

Private Sub lblClose_Click()
    Unload Me
End Sub
