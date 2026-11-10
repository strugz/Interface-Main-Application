VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MsHFlxGd.ocx"
Begin VB.Form frmLabProfileExams 
   BorderStyle     =   0  'None
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11340
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   -90
      Width           =   11325
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
         Left            =   120
         TabIndex        =   9
         Top             =   150
         Width           =   3630
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      TabIndex        =   10
      Top             =   300
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
      Begin VB.Label Label15 
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
         Left            =   150
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   240
         Width           =   4755
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      Enabled         =   0   'False
      ForeColor       =   &H00C0C0C0&
      Height          =   4785
      Left            =   0
      TabIndex        =   13
      Top             =   870
      Width           =   4245
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy Factor and Format from..."
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
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Move &Back"
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
         TabIndex        =   7
         Top             =   4350
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.CommandButton cmdForward 
         Caption         =   "Move &Forward"
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
         TabIndex        =   6
         Top             =   3960
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "Move &Down"
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
         Top             =   3570
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "Move &Up"
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
         TabIndex        =   4
         Top             =   3180
         Visible         =   0   'False
         Width           =   4095
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
         Width           =   3075
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
         Width           =   3075
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
         Left            =   2970
         TabIndex        =   3
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label5 
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
         Left            =   450
         TabIndex        =   15
         Top             =   270
         Width           =   510
      End
      Begin VB.Label Label7 
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
         Left            =   90
         TabIndex        =   14
         Top             =   690
         Width           =   915
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   4785
      Left            =   4230
      TabIndex        =   17
      Top             =   870
      Width           =   7095
      Begin VB.CommandButton cmdAddExam 
         Caption         =   "&Add Exam"
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
         TabIndex        =   22
         Top             =   4260
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtExamID 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   210
         TabIndex        =   21
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
         Left            =   5550
         TabIndex        =   19
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
         Left            =   3990
         TabIndex        =   18
         Top             =   4230
         Visible         =   0   'False
         Width           =   1485
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgExams 
         Height          =   3915
         Left            =   180
         TabIndex        =   20
         Top             =   210
         Width           =   6765
         _ExtentX        =   11933
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

Private Sub cmdAddExam_Click()
Dim tempExamID As String
Dim tempExamName As String
Dim ctr As Integer
Dim examfound As Boolean
examfound = False
With msfgExams
    yClassForms.showSearchCode 3, txtExamID.Text
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
            .Cols = 8
            .TextMatrix(0, 0) = "Code"
            .TextMatrix(0, 1) = "Exam Name"
            .TextMatrix(0, 2) = "Instrument Exam Name"
            .TextMatrix(0, 3) = "Factor"
            .TextMatrix(0, 4) = "DataFormat"
            .TextMatrix(0, 5) = "ProfileID"
            .TextMatrix(0, 6) = "SortNum"
            .TextMatrix(0, 7) = "ParentID"
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
End Sub

Private Sub cmdBack_Click()
Dim tempStr As String
With msfgExams
    .TextMatrix(.Row, 7) = Str(Val(.TextMatrix(.Row, 7)) - 1)
    .TextMatrix(.Row, 1) = String(4 * Val(.TextMatrix(.Row, 7)), " ") & Trim(.TextMatrix(.Row, 1))
    msfgExams_Click
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
            With msfgExams
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
End Sub

Private Sub cmdDown_Click()
Dim tempStr(0 To 7) As String
Dim ctr As Integer
With msfgExams
    For ctr = 0 To 7
        tempStr(ctr) = .TextMatrix(.Row, ctr)
        .TextMatrix(.Row, ctr) = .TextMatrix(.Row + 1, ctr)
        .TextMatrix(.Row + 1, ctr) = tempStr(ctr)
    Next ctr
    .Row = .Row + 1
    .RowSel = .Row
    .ColSel = 4
    msfgExams_Click
    If .RowIsVisible(.Row) = False Then .TopRow = .TopRow + 1
End With
End Sub

Private Sub cmdForward_Click()
Dim tempStr As String
With msfgExams
    .TextMatrix(.Row, 7) = Str(Val(.TextMatrix(.Row, 7)) + 1)
    .TextMatrix(.Row, 1) = String(4 * Val(.TextMatrix(.Row, 7)), " ") & Trim(.TextMatrix(.Row, 1))
    msfgExams_Click
End With
End Sub

Private Sub cmdSave_Click()
Dim ctr As Integer
With msfgExams
    If .Visible Then
        For ctr = 1 To .Rows - 1
            If Trim(.TextMatrix(ctr, 6)) = "NEW" Then
                myConnection.Execute "Insert into tbLabProfileExams (ProfileID, ExamID, SortNum, ParentID, Dilution, DataFormat) " & _
                    "Select '" & .TextMatrix(ctr, 5) & "','" & Trim(.TextMatrix(ctr, 0)) & "'," & ctr & ",'" & Trim(.TextMatrix(ctr, 7)) & "'," & Val(.TextMatrix(ctr, 3)) & ",'" & .TextMatrix(ctr, 4) & "'"
            Else
                myConnection.Execute "Update tbLabProfileExams set SortNum = " & ctr & ", ParentID = '" & Trim(.TextMatrix(ctr, 7)) & "', Dilution = " & Val(.TextMatrix(ctr, 3)) & ", DataFormat = '" & Trim(.TextMatrix(ctr, 4)) & "' where ProfileID = '" & txtProfileID.Text & "' and ExamID = '" & Trim(.TextMatrix(ctr, 0)) & "'"
            End If
        Next ctr
        rstProfile.Requery
        Set msfgExams.DataSource = rstProfile
        MsgBox "Profile Exams has been saved.", vbOKOnly + vbInformation, "Build File"
    End If
End With
End Sub

Private Sub cmdSet_Click()
With msfgExams
    If Len(Trim(txtDilution.Text)) > 0 Then .TextMatrix(.Row, 3) = txtDilution.Text Else .TextMatrix(.Row, 3) = "1"
    .TextMatrix(.Row, 4) = txtDataFormat.Text
End With
End Sub

Private Sub cmdUp_Click()
Dim tempStr(0 To 7) As String
Dim ctr As Integer
With msfgExams
    For ctr = 0 To 7
        tempStr(ctr) = .TextMatrix(.Row, ctr)
        .TextMatrix(.Row, ctr) = .TextMatrix(.Row - 1, ctr)
        .TextMatrix(.Row - 1, ctr) = tempStr(ctr)
    Next ctr
    .Row = .Row - 1
    .RowSel = .Row
    .ColSel = 4
    msfgExams_Click
    If .RowIsVisible(.Row) = False Then .TopRow = .Row
End With
End Sub

Private Sub Form_Load()
msfgExams.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rstProfile = Nothing
    Set rstExam = Nothing
End Sub

Private Sub lblProfileName_Change()
    If Len(Trim(lblProfileName)) > 0 Then
        msfgExams.Visible = True
        cmdUp.Visible = True
        cmdDown.Visible = True
        cmdForward.Visible = True
        cmdBack.Visible = True
        Frame2.Enabled = True
        cmdAddExam.Visible = True
        cmdSave.Visible = True
        txtExamID.Visible = True
        cmdCopy.Visible = True
    Else
        msfgExams.Visible = False
        cmdUp.Visible = False
        cmdDown.Visible = False
        cmdForward.Visible = False
        cmdBack.Visible = False
        Frame2.Enabled = False
        txtDilution.Text = ""
        txtDataFormat.Text = ""
        cmdAddExam.Visible = False
        cmdSave.Visible = False
        txtExamID.Visible = False
        cmdCopy.Visible = False
    End If
End Sub

Private Sub msfgExams_Click()
With msfgExams
    .RowSel = .Row
    txtDilution.Text = .TextMatrix(.Row, 3)
    txtDataFormat.Text = .TextMatrix(.Row, 4)
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
End With
End Sub

Private Sub msfgExams_EnterCell()
msfgExams_Click
End Sub

Private Sub msfgExams_KeyDown(KeyCode As Integer, Shift As Integer)
With msfgExams
If KeyCode = vbKeyDelete And Shift = 1 Then
    If MsgBox("Are you sure you want to remove this item?", vbYesNo + vbQuestion + vbDefaultButton2, "Delete Item") = vbYes Then
        myConnection.Execute "Delete From tbLabProfileExams where ProfileID = '" & txtProfileID.Text & "' and ExamID = '" & .TextMatrix(.Row, 0) & "'"
        Set rstProfile = myConnection.Execute("sp_GerProfileExams '" & txtProfileID.Text & "'")
        If rstProfile.RecordCount > 0 Then
            Set msfgExams.DataSource = rstProfile
            setFlexGrid
            msfgExams_Click
            msfgExams.Visible = True
        Else
            msfgExams.Visible = False
            msfgExams.Clear
            msfgExams.Rows = 2
            msfgExams.Cols = 8
            msfgExams.FixedRows = 1
            cmdUp.Visible = False
            cmdDown.Visible = False
            cmdForward.Visible = False
            cmdBack.Visible = False
            Frame2.Enabled = False
            txtDilution.Text = ""
            txtDataFormat.Text = ""
        End If
    End If
End If
End With
End Sub

Private Sub txtDataFormat_GotFocus()
    highLight
End Sub

Private Sub txtDilution_GotFocus()
    highLight
End Sub

Private Sub txtExamID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdAddExam_Click
If txtExamID.Visible Then txtExamID.SetFocus
End Sub

Private Sub txtExamID_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
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
        Set rstProfile = myConnection.Execute("sp_GerProfileExams '" & txtProfileID.Text & "'")
        If rstProfile.RecordCount > 0 Then
            Set msfgExams.DataSource = rstProfile
            setFlexGrid
            msfgExams_Click
            msfgExams.Visible = True
            msfgExams.Redraw = False
            msfgExams.Redraw = True
        Else
            msfgExams.Visible = False
            msfgExams.Clear
            msfgExams.Rows = 2
            msfgExams.Cols = 8
            msfgExams.FixedRows = 1
            cmdUp.Visible = False
            cmdDown.Visible = False
            cmdForward.Visible = False
            cmdBack.Visible = False
            Frame2.Enabled = False
            txtDilution.Text = ""
            txtDataFormat.Text = ""
            
        End If
    End If
End If
End Sub

Private Sub setFlexGrid()
With msfgExams
    .Height = .CellHeight * 16
    .ColWidth(0) = .Width * 0.1 - (10 * 3)
    .ColWidth(1) = .Width * 0.3 - (30 * 3)
    .ColWidth(2) = .Width * 0.3 - (30 * 3)
    .ColWidth(3) = .Width * 0.15 - (15 * 3)
    .ColWidth(4) = .Width * 0.15 - (15 * 3)
    .ColWidth(5) = 0
    .ColWidth(6) = 0
    .ColWidth(7) = 0
    .ColSel = 4
End With
End Sub
