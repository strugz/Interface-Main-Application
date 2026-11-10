VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmClassifications 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Normal Value"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6720
      TabIndex        =   0
      Top             =   -120
      Width           =   6135
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Build File - [Classifications]"
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
         TabIndex        =   1
         Top             =   150
         Width           =   2565
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   4815
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   -60
      Width           =   6135
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   4515
         Left            =   120
         TabIndex        =   3
         Top             =   180
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   7964
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   12615680
         ForeColorFixed  =   16777215
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   1995
      Index           =   1
      Left            =   60
      TabIndex        =   4
      Top             =   4680
      Width           =   6135
      Begin VB.CommandButton cmdSave 
         Caption         =   "Add"
         Height          =   435
         Left            =   4740
         TabIndex        =   9
         Top             =   1410
         Width           =   1215
      End
      Begin VB.ComboBox cboStatus 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1020
         Width           =   1485
      End
      Begin VB.TextBox txtProfileName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1350
         TabIndex        =   7
         Top             =   630
         Width           =   4635
      End
      Begin VB.TextBox txtProfileID 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   210
         Width           =   1425
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   435
         Left            =   3480
         TabIndex        =   5
         Top             =   1410
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Clss ID:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   690
         TabIndex        =   12
         Top             =   270
         Width           =   540
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Class Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   270
         TabIndex        =   11
         Top             =   630
         Width           =   900
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   690
         TabIndex        =   10
         Top             =   1020
         Width           =   510
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   675
      Index           =   2
      Left            =   60
      TabIndex        =   13
      Top             =   6690
      Width           =   6135
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print..."
         Height          =   435
         Left            =   3510
         TabIndex        =   15
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   435
         Left            =   4800
         TabIndex        =   14
         Top             =   180
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmClassifications"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstProfiles As ADODB.Recordset
Private rstProfileID As ADODB.Recordset

Private Sub cmdCancel_Click()
    cmdCancel.Visible = False
    txtProfileID.Text = Val(rstProfileID.Fields("ClassID").Value) + 1
    cboStatus.ListIndex = 0
    txtProfileName.Text = ""
    cmdSave.Caption = "Add"
    msfgResult.Enabled = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
Me.MousePointer = 11
    If UCase(Dir(App.Path & "\Reports", vbDirectory)) = "REPORTS" Then
        OpenMainReport App.Path & "\reports\" & "Classifications.rpt", CStr("1"), InstrumentID
    ElseIf UCase(Dir(Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports", vbDirectory)) = "REPORTS" Then
        OpenMainReport Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports\" & "Classifications.rpt", CStr("1"), InstrumentID
    End If
    ShowReportViewer False, "Result Print Preview", False
Me.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
If cmdSave.Caption = "Add" Then
    If Len(Trim(txtProfileName.Text)) > 0 Then
        myConnection.Execute "Insert into tbLabClassifications (classID,ClassName,Active,InstrumentID) select '" & txtProfileID.Text & "','" & Replace(txtProfileName, "'", "''") & "'," & cboStatus.Text & ",'" & InstrumentID & "'"
        rstProfiles.Requery
        If rstProfiles.RecordCount > 0 Then
            Set msfgResult.DataSource = rstProfiles
            rstProfileID.Requery
            txtProfileName.Text = ""
            cboStatus.ListIndex = 0
            txtProfileID.Text = Val(rstProfileID.Fields("ClassID").Value) + 1
            msfgResult.Visible = True
        End If
    End If

Else
    myConnection.Execute "Update tbLabClassifications set ClassName = '" & Replace(txtProfileName.Text, "'", "''") & "', Active = " & cboStatus.Text & " where classID = '" & txtProfileID.Text & "'"
    rstProfiles.Requery
    If rstProfiles.RecordCount > 0 Then
        Set msfgResult.DataSource = rstProfiles
        rstProfileID.Requery
        txtProfileName.Text = ""
        cboStatus.ListIndex = 0
        txtProfileID.Text = Val(rstProfileID.Fields("ClassID").Value) + 1
        msfgResult.Visible = True
        cmdCancel.Visible = False
        msfgResult.Enabled = True
        cmdSave.Caption = "Add"
    End If
End If
End Sub

Private Sub Form_Load()
setFormColors Me
    Set rstProfiles = myConnection.Execute("Select * from tbLabClassifications where InstrumentID = '" & InstrumentID & "'")
    Set rstProfileID = myConnection.Execute("Select max(classid) as classID from tbLabClassifications where InstrumentID = '" & InstrumentID & "'")
    If rstProfiles.RecordCount > 0 Then
        Set msfgResult.DataSource = rstProfiles
        setFlexGrid
        txtProfileID.Text = Val(rstProfileID.Fields("ClassID").Value) + 1
        msfgResult.Visible = True
    Else
        msfgResult.Visible = False
        txtProfileID.Text = 1
    End If
    cboStatus.ListIndex = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set rstProfiles = Nothing
End Sub

Private Sub setFlexGrid()
With msfgResult
    .Height = .CellHeight * 16
    .ColWidth(0) = .Width * 0.2 - (20 * 3)
    .ColWidth(1) = .Width * 0.7 - (50 * 3)
    .ColWidth(2) = .Width * 0.1 - (10 * 3)
    .ColWidth(3) = 0
    .ColSel = 2
End With
End Sub

Private Sub msfgResult_DblClick()
With msfgResult
    msfgResult.Enabled = False
    cmdSave.Caption = "Update"
    cmdCancel.Visible = True
    txtProfileID.Text = .TextMatrix(.Row, 0)
    txtProfileName.Text = .TextMatrix(.Row, 1)
    cboStatus.ListIndex = IIf((.TextMatrix(.Row, 2) = "True"), 0, 1)
End With
End Sub

Private Sub txtProfileName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

