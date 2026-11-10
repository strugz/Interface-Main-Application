VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmRptViewer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report Viewer"
   ClientHeight    =   3645
   ClientLeft      =   4365
   ClientTop       =   3495
   ClientWidth     =   5205
   Icon            =   "frmRptViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   3405
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   5025
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
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmRptViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private yRefNum As String
Private yTag As String
Dim blnDirectPrint As Boolean

Private Sub CRViewer1_CloseButtonClicked(UseDefault As Boolean)
    Unload Me
End Sub

Private Sub CRViewer1_ExportButtonClicked(UseDefault As Boolean)
Dim strFormat As Integer
Dim FileExt As String
Dim Filename As String
Dim arrTemp() As String
Dim tempStr As String
Dim tempTitle As String
UseDefault = False
tempTitle = Report.ReportTitle
Select Case UCase(Me.Caption)
    Case "LOGBOOK PREVIEW"
        Filename = "LOGBOOK" & "_" & tempTitle
    Case "WORKLIST PREVIEW"
        Filename = "WORKLIST" & "_" & tempTitle
    Case "RESULT PRINT PREVIEW"
        Filename = "RESULT" & "_" & tempTitle
End Select
Report.ReportTitle = "Sheet1"
Filename = Replace(Filename, "/", "-")
Filename = Replace(Filename, ":", "")
tempStr = ShowExportReport(Filename)
If InStr(1, tempStr, "|") Then
    arrTemp = Split(tempStr, "|")
    Select Case arrTemp(0)
        Case "Comma Separated Values (CSV)"
            strFormat = 5
            FileExt = ".csv"
        Case "Excel (XLS)"
            strFormat = 29
            FileExt = ".xls"
        Case "Rich Text Format (RTF)"
            strFormat = 4
            FileExt = ".rtf"
        Case "HTML"
            strFormat = 32
            FileExt = ".html"
        Case "Word for Windows (DOC)"
            strFormat = 14
            FileExt = ".doc"
        Case "Text File"
            strFormat = 8
            FileExt = ".txt"
    End Select
    Report.ExportOptions.DestinationType = crEDTDiskFile
    Report.ExportOptions.FormatType = strFormat
    Report.ExportOptions.DiskFileName = arrTemp(1) & "\" & arrTemp(2) & FileExt
    Report.Export False
    Report.ReportTitle = tempTitle
    MsgBox "Report saved as " & arrTemp(1) & "\" & arrTemp(2) & FileExt
End If
End Sub

Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
    CRViewer1.PrintReport
    If Me.Caption = "Result Print Preview" Then
        myConnection.Execute "Update tbLabInstrumentResults set Printed = TRUE, Status = '2' where refnum ='" & yRefNum & "'"
        myConnection.Execute "Update tbPatientMaster set Printed = TRUE, PrintDate = #" & Format(Now, "mm/dd/yyyy HH:mm") & "# where refnum ='" & yRefNum & "'"
    End If
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
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
    DoEvents
    Do While CRViewer1.IsBusy
    Me.MousePointer = 11
    Loop
    Me.MousePointer = 0
    CRViewer1.Zoom 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Set Report = Nothing
End Sub

Public Property Let directPrint(ByVal blnNewValue As Boolean)
   blnDirectPrint = blnNewValue
End Property

Public Property Let wRefNum(ByVal strNewValue As String)
   yRefNum = strNewValue
End Property

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuPrint_Click()
    CRViewer1.PrintReport
    Unload Me
End Sub

Public Property Let wTag(ByVal vNewValue As String)
    yTag = vNewValue
End Property
