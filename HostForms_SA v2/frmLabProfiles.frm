VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmLabProfiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profiles"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8820
      TabIndex        =   0
      Top             =   1830
      Width           =   3195
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Build File - [Laboratory Profiles]"
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
         Width           =   3075
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   8055
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   3915
         Left            =   120
         TabIndex        =   3
         Top             =   180
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   6906
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
      BackColor       =   &H006B4B01&
      Height          =   1935
      Index           =   1
      Left            =   30
      TabIndex        =   6
      Top             =   4260
      Width           =   8055
      Begin VB.CommandButton cmdTemplate 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4890
         TabIndex        =   18
         Top             =   1020
         Width           =   345
      End
      Begin VB.TextBox txtTemplate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   990
         Width           =   3525
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   405
         Left            =   5400
         TabIndex        =   14
         Top             =   1410
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtProfileID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   210
         Width           =   1425
      End
      Begin VB.TextBox txtProfileName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1350
         TabIndex        =   12
         Top             =   630
         Width           =   4635
      End
      Begin VB.ComboBox cboStatus 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmLabProfiles.frx":0000
         Left            =   1320
         List            =   "frmLabProfiles.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1350
         Width           =   1485
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Add"
         Height          =   405
         Left            =   6660
         TabIndex        =   10
         Top             =   1410
         Width           =   1215
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Template:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   510
         TabIndex        =   17
         Top             =   990
         Width           =   690
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   690
         TabIndex        =   9
         Top             =   1350
         Width           =   510
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Profile Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   270
         TabIndex        =   8
         Top             =   630
         Width           =   945
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Profile ID:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   540
         TabIndex        =   7
         Top             =   270
         Width           =   675
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   675
      Index           =   2
      Left            =   30
      TabIndex        =   4
      Top             =   6210
      Width           =   8055
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5430
         TabIndex        =   15
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6690
         TabIndex        =   5
         Top             =   180
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmLabProfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstProfiles As ADODB.Recordset
Private rstProfileID As ADODB.Recordset

Private Sub cmdCancel_Click()
    cmdCancel.Visible = False
    txtProfileID.Text = Val(rstProfileID.Fields("ProfileID").Value) + 1
    cboStatus.ListIndex = 0
    txtProfileName.Text = ""
    cmdSave.Caption = "Add"
    msfgResult.Enabled = True
    txtTemplate.Text = ""
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
Me.MousePointer = 11
    If UCase(Dir(App.Path & "\Reports", vbDirectory)) = "REPORTS" Then
        OpenMainReport App.Path & "\reports\" & "Profiles.rpt", CStr("1"), InstrumentID
    ElseIf UCase(Dir(Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports", vbDirectory)) = "REPORTS" Then
        OpenMainReport Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports\" & "Profiles.rpt", CStr("1"), InstrumentID
    End If
    ShowReportViewer False, "Result Print Preview", False
Me.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
If cmdSave.Caption = "Add" Then
    If Len(Trim(txtProfileName.Text)) > 0 Then
        myConnection.Execute "Insert into tbLabProfiles (profileID,ProfileName,Status,InstrumentID, Template) select '" & txtProfileID.Text & "','" & txtProfileName & "','" & IIf((cboStatus.Text = "TRUE"), "1'", "0'") & ",'" & InstrumentID & "','" & txtTemplate.Text & "'"
        rstProfiles.Requery
        If rstProfiles.RecordCount > 0 Then
            Set msfgResult.DataSource = rstProfiles
            rstProfileID.Requery
            txtProfileName.Text = ""
            txtTemplate.Text = ""
            cboStatus.ListIndex = 0
            txtProfileID.Text = Val(rstProfileID.Fields("ProfileID").Value) + 1
            msfgResult.Visible = True
        End If
    End If
Else
    myConnection.Execute "Update tbLabProfiles set profileName = '" & Replace(txtProfileName.Text, "'", "''") & "', Status = '" & IIf((cboStatus.Text = "TRUE"), "1'", "0'") & ", Template = '" & txtTemplate.Text & "' where ProfileID = '" & txtProfileID.Text & "' and InstrumentID = '" & InstrumentID & "'"
    rstProfiles.Requery
    If rstProfiles.RecordCount > 0 Then
        Set msfgResult.DataSource = rstProfiles
        rstProfileID.Requery
        txtProfileName.Text = ""
        txtTemplate.Text = ""
        cboStatus.ListIndex = 0
        txtProfileID.Text = Val(rstProfileID.Fields("ProfileID").Value) + 1
        msfgResult.Visible = True
        cmdCancel.Visible = False
        msfgResult.Enabled = True
        cmdSave.Caption = "Add"
    End If
End If
End Sub

Private Sub cmdTemplate_Click()
    frmTemplate.Show 1
    txtTemplate = TemplateValue
End Sub

Private Sub Form_Load()
setFormColors Me
    Set rstProfiles = myConnection.Execute("Select * from tblabprofiles where InstrumentID = '" & InstrumentID & "'")
    If DBServerType = "LOCAL DATABASE" Then
        Set rstProfileID = myConnection.Execute("Select max(cint(profileid)) as profileID from tblabprofiles where InstrumentID = '" & InstrumentID & "'")
    ElseIf DBServerType = "MYSQL SERVER" Then
        Set rstProfileID = myConnection.Execute("Select max(cast(profileid as signed)) ""profileID"" from tblabprofiles where InstrumentID = '" & InstrumentID & "'")
    ElseIf DBServerType = "MSSQL SERVER" Then
        Set rstProfileID = myConnection.Execute("Select max(cast(profileid as int)) ""profileID"" from tblabprofiles where InstrumentID = '" & InstrumentID & "'")
    End If
    If rstProfiles.RecordCount > 0 Then
        Set msfgResult.DataSource = rstProfiles
        setFlexGrid
        txtProfileID.Text = Val(rstProfileID.Fields("ProfileID").Value) + 1
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
    .ColWidth(1) = .Width * 0.5 - (50 * 3)
    .ColWidth(2) = .Width * 0.1 - (10 * 3)
    .ColWidth(3) = 0
    .ColWidth(4) = .Width * 0.2 - (10 * 3)
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
    cboStatus.ListIndex = IIf((.TextMatrix(.Row, 2) = 1), 0, 1)
    txtTemplate.Text = .TextMatrix(.Row, 4)
End With
End Sub

Private Sub msfgResult_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 1 And KeyCode = vbKeyDelete Then
    If MsgBox("Are you sure you want to delete this profile?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Item") = vbYes Then
        
    End If
End If
End Sub

Private Sub txtProfileName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtTemplate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        txtTemplate.Text = ""
    End If
End Sub
