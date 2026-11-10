VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmLabProfileExams 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profile Exams"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11445
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   11445
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   11910
      TabIndex        =   10
      Top             =   3510
      Width           =   4005
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Build File - [Laboratory Profile Exams]"
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
         Left            =   90
         TabIndex        =   11
         Top             =   180
         Width           =   3630
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   675
      Index           =   0
      Left            =   30
      TabIndex        =   12
      Top             =   -60
      Width           =   11325
      Begin VB.TextBox txtProfileID 
         Alignment       =   1  'Right Justify
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
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   0
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackColor       =   &H00FCE2C7&
         BackStyle       =   0  'Transparent
         Caption         =   "Profile Exam:"
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
         Left            =   150
         TabIndex        =   14
         Top             =   270
         Width           =   930
      End
      Begin VB.Label lblProfileName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   4755
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   4785
      Index           =   1
      Left            =   30
      TabIndex        =   15
      Top             =   630
      Width           =   3345
      Begin VB.TextBox txtGroup 
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
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1020
         Width           =   2175
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy Format from..."
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
         Left            =   90
         TabIndex        =   5
         Top             =   1470
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   330
         Picture         =   "frmLabProfileExams.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2880
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.CommandButton cmdForward 
         Caption         =   "&Forward"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   2130
         Picture         =   "frmLabProfileExams.frx":1982
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2880
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "&Down"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   1230
         Picture         =   "frmLabProfileExams.frx":3304
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3750
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "&Up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   1230
         Picture         =   "frmLabProfileExams.frx":4C86
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2010
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtDilution 
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
         Left            =   1050
         MaxLength       =   8
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtDataFormat 
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
         Left            =   1050
         MaxLength       =   8
         TabIndex        =   2
         Top             =   630
         Width           =   2175
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "Se&t"
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
         Left            =   2070
         TabIndex        =   4
         Top             =   1470
         Width           =   1125
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Group:"
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
         TabIndex        =   24
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Factor:"
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
         Left            =   450
         TabIndex        =   17
         Top             =   270
         Width           =   510
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Format:"
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
         Left            =   90
         TabIndex        =   16
         Top             =   690
         Width           =   915
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   4785
      Index           =   2
      Left            =   3420
      TabIndex        =   18
      Top             =   630
      Width           =   7995
      Begin VB.CommandButton cmdAddExam 
         Caption         =   "&Add Exam"
         Default         =   -1  'True
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
         Left            =   2610
         TabIndex        =   23
         Top             =   4260
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtExamID 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   210
         TabIndex        =   22
         Top             =   4230
         Visible         =   0   'False
         Width           =   2385
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
         Height          =   375
         Left            =   6480
         TabIndex        =   20
         Top             =   4230
         Width           =   1395
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   19
         Top             =   4230
         Visible         =   0   'False
         Width           =   1485
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   3915
         Left            =   180
         TabIndex        =   21
         Top             =   210
         Width           =   7725
         _ExtentX        =   13626
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
End
Attribute VB_Name = "frmLabProfileExams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstProfile As ADODB.Recordset
Private rstExam As ADODB.Recordset
Private yClassForms As New ClassForms
Private BatchUpdate As Boolean

Private Sub cmdAddExam_Click()
Dim tempExamID As String
Dim tempExamName As String
Dim ctr As Integer
Dim examfound As Boolean
examfound = False
With msfgResult
    yClassForms.showSearchCode 3, txtExamID.Text, False
    tempExamID = yClassForms.LastsearchID
    tempExamName = yClassForms.LastSearchName
    If Len(Trim(tempExamName)) > 0 Then
        If rstProfile.RecordCount > 0 Then
            For ctr = 1 To .Rows - 1
                If Trim(.TextMatrix(ctr, 0)) = Trim(tempExamID) Then
                    MsgBox "Exam Already Exists! Press OK to continue", vbOKOnly + vbInformation, "Build Files"
                    examfound = True
                End If
            Next ctr
        Else
            .Visible = True
            .Cols = 9
            .TextMatrix(0, 0) = "Code"
            .TextMatrix(0, 1) = "Exam Name"
            .TextMatrix(0, 2) = "Inst. Exam Name"
            .TextMatrix(0, 3) = "Factor"
            .TextMatrix(0, 4) = "DataFormat"
            .TextMatrix(0, 5) = "ProfileID"
            .TextMatrix(0, 6) = "SortNum"
            .TextMatrix(0, 7) = "ParentID"
            .TextMatrix(0, 8) = "Group"
            For ctr = 1 To .Rows - 1
                If Trim(.TextMatrix(ctr, 0)) = Trim(tempExamID) Then
                    MsgBox "Exam Already Exists! Press OK to continue", vbOKOnly + vbInformation, "Build Files"
                    examfound = True
                End If
            Next ctr
        End If
        If Not examfound Then
            Set rstExam = myConnection.Execute("Select * from tbcolabexam where examid = '" & tempExamID & "'")
            If rstExam.RecordCount > 0 Then
                lblProfileName_Change
                .Redraw = False
                If Trim(.TextMatrix(.Rows - 1, 0)) <> "" Then .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = tempExamID
                .TextMatrix(.Rows - 1, 1) = rstExam.Fields("DisplayName").Value
                .TextMatrix(.Rows - 1, 2) = rstExam.Fields("InstExamName").Value
                .TextMatrix(.Rows - 1, 3) = "1"
                .TextMatrix(.Rows - 1, 4) = ""
                .TextMatrix(.Rows - 1, 5) = txtProfileID.Text
                .TextMatrix(.Rows - 1, 6) = "NEW"
                .TextMatrix(.Rows - 1, 7) = "0"
                setFlexGrid
                .Redraw = True
            End If
        End If
    Else
    End If
    txtExamID.Text = ""
    If txtExamID.Visible And txtExamID.Enabled Then txtExamID.SetFocus
End With
If txtExamID.Visible Then txtExamID.SetFocus
msfgResult.TopRow = msfgResult.Rows - 1
msfgResult.Row = msfgResult.Rows - 1
End Sub

Private Sub cmdBack_Click()
Dim tempStr As String
With msfgResult
    .TextMatrix(.Row, 7) = Str(Val(.TextMatrix(.Row, 7)) - 1)
    .TextMatrix(.Row, 1) = String(4 * Val(.TextMatrix(.Row, 7)), " ") & Trim(.TextMatrix(.Row, 1))
    msfgResult_Click
End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
Dim tempProfileID As String
Dim rstTemp As ADODB.Recordset
Dim ctr As Integer
    yClassForms.showSearchCode 1, ""
    tempProfileID = yClassForms.LastsearchID
    If Len(Trim(tempProfileID)) > 0 Then
        Set rstTemp = myConnection.Execute("Select * from tbLabProfileExams where profileid = '" & tempProfileID & "'")
        If rstTemp.RecordCount > 0 Then
            With msfgResult
                For ctr = 1 To .Rows - 1
                    rstTemp.MoveFirst
                    rstTemp.Find "ExamID = '" & .TextMatrix(ctr, 0) & "'"
                    If Not rstTemp.EOF Then
                        .TextMatrix(ctr, 3) = rstTemp.Fields("Dilution").Value
                        .TextMatrix(ctr, 4) = rstTemp.Fields("DataFormat").Value
                    End If
                Next ctr
            End With
        End If
    End If
setGroupOrder
End Sub

Private Sub cmdDown_Click()
Dim tempStr(0 To 8) As String
Dim ctr As Integer
With msfgResult
    For ctr = 0 To 8
        tempStr(ctr) = .TextMatrix(.Row, ctr)
        .TextMatrix(.Row, ctr) = .TextMatrix(.Row + 1, ctr)
        .TextMatrix(.Row + 1, ctr) = tempStr(ctr)
    Next ctr
    .Row = .Row + 1
    .RowSel = .Row
    .ColSel = 8
    msfgResult_Click
    If .RowIsVisible(.Row) = False Then .TopRow = .TopRow + 1
End With
setGroupOrder
End Sub

Private Sub cmdForward_Click()
Dim tempStr As String
With msfgResult
    .TextMatrix(.Row, 7) = Str(Val(.TextMatrix(.Row, 7)) + 1)
    .TextMatrix(.Row, 1) = String(4 * Val(.TextMatrix(.Row, 7)), " ") & Trim(.TextMatrix(.Row, 1))
    msfgResult_Click
End With
End Sub

Private Sub cmdSave_Click()
Dim ctr As Integer
With msfgResult
    If .Visible Then
        For ctr = 1 To .Rows - 1
            If Trim(.TextMatrix(ctr, 6)) = "NEW" Then
                .Cols = 10
                myConnection.Execute "Insert into tbLabProfileExams (ProfileID, ExamID, SortNum, ParentID, Dilution, DataFormat, InstrumentID, [Group], GroupOrder) " & _
                    "Select '" & .TextMatrix(ctr, 5) & "','" & Trim(.TextMatrix(ctr, 0)) & "'," & ctr & ",'" & Trim(.TextMatrix(ctr, 7)) & "'," & Val(.TextMatrix(ctr, 3)) & ",'" & .TextMatrix(ctr, 4) & "','" & InstrumentID & "','" & Trim(.TextMatrix(ctr, 8)) & "'," & Val(.TextMatrix(ctr, 9))
            Else
                myConnection.Execute "Update tbLabProfileExams set SortNum = " & ctr & ", ParentID = '" & Trim(.TextMatrix(ctr, 7)) & "', Dilution = " & Val(.TextMatrix(ctr, 3)) & ", DataFormat = '" & Trim(.TextMatrix(ctr, 4)) & "', [Group] = '" & Replace(.TextMatrix(ctr, 8), "'", "''") & "', GroupOrder = " & .TextMatrix(ctr, 9) & " where ProfileID = '" & txtProfileID.Text & "' and ExamID = '" & Trim(.TextMatrix(ctr, 0)) & "'"
            End If
        Next ctr
        rstProfile.Requery
        Set msfgResult.DataSource = rstProfile
        MsgBox "Profile Exams has been saved.", vbOKOnly + vbInformation, "Build File"
    End If
End With
End Sub

Private Sub cmdSet_Click()
Dim tmpRow As Integer
Dim ctr As Integer
With msfgResult
If BatchUpdate Then
    If .Row > .RowSel Then
        tmpRow = .Row
        .Row = .RowSel
        .RowSel = tmpRow
    End If
    For ctr = .Row To .RowSel
        If txtDataFormat.BackColor = &H80000005 Then .TextMatrix(ctr, 4) = txtDataFormat.Text
        If txtDilution.BackColor = &H80000005 Then
            If Len(Trim(txtDilution.Text)) > 0 Then .TextMatrix(ctr, 3) = txtDilution.Text Else .TextMatrix(ctr, 3) = "1"
        End If
        If txtGroup.BackColor = &H80000005 Then .TextMatrix(ctr, 8) = txtGroup.Text
    Next ctr
Else
    If Len(Trim(txtDilution.Text)) > 0 Then .TextMatrix(.Row, 3) = txtDilution.Text Else .TextMatrix(.Row, 3) = "1"
    .TextMatrix(.Row, 4) = txtDataFormat.Text
    .TextMatrix(.Row, 8) = txtGroup.Text
End If
End With
setGroupOrder
End Sub

Private Sub cmdUp_Click()
Dim tempStr(0 To 8) As String
Dim ctr As Integer
With msfgResult
    For ctr = 0 To 8
        tempStr(ctr) = .TextMatrix(.Row, ctr)
        .TextMatrix(.Row, ctr) = .TextMatrix(.Row - 1, ctr)
        .TextMatrix(.Row - 1, ctr) = tempStr(ctr)
    Next ctr
    .Row = .Row - 1
    .RowSel = .Row
    .ColSel = 8
    msfgResult_Click
    If .RowIsVisible(.Row) = False Then .TopRow = .Row
End With
setGroupOrder
End Sub

Private Sub Form_Load()
setFormColors Me
msfgResult.Visible = False
BatchUpdate = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rstProfile = Nothing
    Set rstExam = Nothing
End Sub

Private Sub lblProfileName_Change()
    If Len(Trim(lblProfileName)) > 0 Then
        msfgResult.Visible = True
        cmdUp.Visible = True
        cmdDown.Visible = True
        cmdForward.Visible = True
        cmdBack.Visible = True
        fraBody(1).Enabled = True
        cmdAddExam.Visible = True
        cmdSave.Visible = True
        txtExamID.Visible = True
        cmdCopy.Visible = True
    Else
        msfgResult.Visible = False
        cmdUp.Visible = False
        cmdDown.Visible = False
        cmdForward.Visible = False
        cmdBack.Visible = False
        fraBody(1).Enabled = False
        txtDilution.Text = ""
        txtDataFormat.Text = ""
        cmdAddExam.Visible = False
        cmdSave.Visible = False
        txtExamID.Visible = False
        cmdCopy.Visible = False
    End If
End Sub

Private Sub msfgResult_Click()
With msfgResult
    If .RowSel <> .Row Then
        BatchUpdate = True
        txtDilution.SetFocus
    Else
        BatchUpdate = False
        txtGroup.BackColor = &H80000005
        txtDilution.BackColor = &H80000005
        txtDataFormat.BackColor = &H80000005
    End If
    txtDilution.Text = .TextMatrix(.Row, 3)
    txtDataFormat.Text = .TextMatrix(.Row, 4)
    txtGroup.Text = .TextMatrix(.Row, 8)
    If BatchUpdate Then
        cmdUp.Enabled = False
        cmdDown.Enabled = False
        cmdForward.Enabled = False
        cmdBack.Enabled = False
    Else
        If .Row = .Rows - 1 Then
            cmdDown.Enabled = False
        Else
            cmdDown.Enabled = True
        End If
        If Val(.TextMatrix(.Row, 7)) > 3 Then
            cmdForward.Enabled = False
        Else
            cmdForward.Enabled = True
        End If
        If Val(.TextMatrix(.Row, 7)) = 0 Then
            cmdBack.Enabled = False
        Else
            cmdBack.Enabled = True
        End If
        If .Row = 1 Then
            cmdUp.Enabled = False
        Else
            cmdUp.Enabled = True
        End If
    End If
End With
End Sub

Private Sub msfgResult_EnterCell()
msfgResult_Click
End Sub

Private Sub msfgResult_KeyDown(KeyCode As Integer, Shift As Integer)
With msfgResult
If KeyCode = vbKeyDelete And Shift = 1 Then
    If MsgBox("Are you sure you want to remove this item?", vbYesNo + vbQuestion + vbDefaultButton2, "Delete Item") = vbYes Then
        myConnection.Execute "Delete From tbLabProfileExams where ProfileID = '" & txtProfileID.Text & "' and ExamID = '" & .TextMatrix(.Row, 0) & "'"
        Set rstProfile = myConnection.Execute("sp_GetProfileExams '" & txtProfileID.Text & "','" & InstrumentID & "'")
        If rstProfile.RecordCount > 0 Then
            Set msfgResult.DataSource = rstProfile
            setFlexGrid
            msfgResult_Click
            msfgResult.Visible = True
        Else
            msfgResult.Visible = False
            msfgResult.Clear
            msfgResult.Rows = 2
            msfgResult.Cols = 9
            msfgResult.FixedRows = 1
            cmdUp.Visible = False
            cmdDown.Visible = False
            cmdForward.Visible = False
            cmdBack.Visible = False
            fraBody(1).Enabled = False
            txtDilution.Text = ""
            txtDataFormat.Text = ""
        End If
    End If
End If
End With
End Sub

Private Sub txtDataFormat_GotFocus()
    highLight
    If BatchUpdate Then
        txtDataFormat.BackColor = &HC0E0FF
        txtGroup.BackColor = &H8000000F
        txtDilution.BackColor = &H8000000F
    Else
        txtDataFormat.BackColor = &H80000005
    End If
End Sub

Private Sub txtDataFormat_LostFocus()
    txtDataFormat.BackColor = &H80000005
End Sub

Private Sub txtDilution_GotFocus()
    highLight
    If BatchUpdate Then
        txtDilution.BackColor = &HC0E0FF
        txtGroup.BackColor = &H8000000F
        txtDataFormat.BackColor = &H8000000F
    Else
        txtDilution.BackColor = &H80000005
    End If
End Sub

Private Sub txtDilution_LostFocus()
    txtDilution.BackColor = &H80000005
End Sub

Private Sub txtExamID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdAddExam_Click
If txtExamID.Visible Then txtExamID.SetFocus
End Sub

Private Sub txtExamID_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtGroup_GotFocus()
    highLight
    If BatchUpdate Then
        txtGroup.BackColor = &HC0E0FF
        txtDilution.BackColor = &H8000000F
        txtDataFormat.BackColor = &H8000000F
    Else
        txtGroup.BackColor = &H80000005
    End If
End Sub

Private Sub txtGroup_LostFocus()
    txtGroup.BackColor = &H80000005
End Sub

Private Sub txtProfileID_Change()
    lblProfileName.Caption = ""
End Sub

Private Sub txtProfileID_GotFocus()
    highLight
End Sub

Private Sub txtProfileID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    yClassForms.showSearchCode 1, txtProfileID
    txtProfileID.Text = yClassForms.LastsearchID
    lblProfileName.Caption = yClassForms.LastSearchName
    If Len(Trim(lblProfileName.Caption)) > 0 Then
        Set rstProfile = myConnection.Execute("sp_GetProfileExams '" & txtProfileID.Text & "','" & InstrumentID & "'")
        If rstProfile.RecordCount > 0 Then
            Set msfgResult.DataSource = rstProfile
            setFlexGrid
            msfgResult_Click
            msfgResult.Visible = True
            msfgResult.Redraw = False
            msfgResult.Redraw = True
            setGroupOrder
        Else
            msfgResult.Visible = False
            msfgResult.Clear
            msfgResult.Rows = 2
            msfgResult.Cols = 9
            msfgResult.FixedRows = 1
            cmdUp.Visible = False
            cmdDown.Visible = False
            cmdForward.Visible = False
            cmdBack.Visible = False
            fraBody(1).Enabled = False
            txtDilution.Text = ""
            txtDataFormat.Text = ""
            
        End If
    End If
End If
End Sub

Private Sub setFlexGrid()
With msfgResult
    .Height = .CellHeight * 16
    .ColWidth(0) = .Width * 0.1 - (10 * 3)
    .ColWidth(1) = .Width * 0.2 - (20 * 3)
    .ColWidth(2) = .Width * 0.2 - (20 * 3)
    .ColWidth(3) = .Width * 0.1 - (10 * 3)
    .ColWidth(4) = .Width * 0.1 - (10 * 3)
    .ColWidth(5) = 0
    .ColWidth(6) = 0
    .ColWidth(7) = 0
    .ColWidth(8) = .Width * 0.2 - (20 * 3)
    .ColWidth(9) = .Width * 0.1 - (10 * 3)
    .ColSel = 8
    
    .TextMatrix(0, 0) = "Code"
    .TextMatrix(0, 1) = "Exam Name"
    .TextMatrix(0, 2) = "Inst. Exam Name"
    .TextMatrix(0, 3) = "Factor"
    .TextMatrix(0, 4) = "DataFormat"
    .TextMatrix(0, 5) = "ProfileID"
    .TextMatrix(0, 6) = "SortNum"
    .TextMatrix(0, 7) = "ParentID"
    .TextMatrix(0, 8) = "Group"
End With
End Sub

Private Sub setGroupOrder()
Dim ctr As Integer
Dim ctr2 As Integer
Dim argInc As Boolean
argInc = False
With msfgResult
    ctr2 = 1
    If .Cols = 10 Then .Cols = 11
    For ctr = 1 To .Rows - 1
        If .Rows - 1 > 1 Then
            If ctr = .Rows - 1 Then
                If argInc Then .TextMatrix(ctr, 9) = ctr2 + 1 Else .TextMatrix(ctr, 9) = ctr2
            Else
                If .TextMatrix(ctr, 8) = .TextMatrix(ctr + 1, 8) Then
                    If argInc Then .TextMatrix(ctr, 9) = ctr2 + 1 Else .TextMatrix(ctr, 9) = ctr2
                    argInc = False
                Else
                    If argInc Then .TextMatrix(ctr, 9) = ctr2 + 1 Else .TextMatrix(ctr, 9) = ctr2
                    argInc = True
                End If
            End If
        Else
            .TextMatrix(ctr, 8) = ctr2
        End If
        ctr2 = Val(.TextMatrix(ctr, 9))
    Next ctr
End With
End Sub
