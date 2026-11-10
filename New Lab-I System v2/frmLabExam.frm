VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLabExam 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   5670
   ClientLeft      =   1770
   ClientTop       =   1710
   ClientWidth     =   11325
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   -90
      Width           =   11325
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Build File - [Laboratory Exam]"
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
         TabIndex        =   10
         Top             =   150
         Width           =   2850
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H00C0C0C0&
      Height          =   5355
      Left            =   0
      TabIndex        =   11
      Top             =   300
      Width           =   11325
      Begin VB.TextBox txtUnit 
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
         Left            =   600
         MaxLength       =   8
         TabIndex        =   22
         Top             =   1530
         Width           =   3015
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   435
         Left            =   570
         TabIndex        =   20
         Top             =   3660
         Visible         =   0   'False
         Width           =   1485
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgExams 
         Height          =   4455
         Left            =   3810
         TabIndex        =   19
         Top             =   330
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   7858
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
      Begin VB.CommandButton cmdAddSave 
         Caption         =   "Add"
         Height          =   435
         Left            =   2100
         TabIndex        =   7
         Top             =   3660
         Width           =   1485
      End
      Begin VB.TextBox txtResultCode 
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
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   5
         Top             =   2670
         Width           =   2265
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   375
         Left            =   9720
         TabIndex        =   8
         Top             =   4860
         Width           =   1395
      End
      Begin VB.TextBox txtInstrumentExamName 
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
         Left            =   1830
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1110
         Width           =   1755
      End
      Begin VB.TextBox txtLISExamCode 
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
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   4
         Top             =   2280
         Width           =   2265
      End
      Begin VB.TextBox txtDisplayName 
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
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1890
         Width           =   2265
      End
      Begin VB.TextBox txtRefNum 
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
         Left            =   1830
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Width           =   1755
      End
      Begin VB.TextBox txtInstrumentID 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   1170
         MaxLength       =   15
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "HMX"
         Top             =   330
         Width           =   2415
      End
      Begin VB.ComboBox cboActive 
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
         ItemData        =   "frmLabExam.frx":0000
         Left            =   750
         List            =   "frmLabExam.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3060
         Width           =   2805
      End
      Begin VB.Label lblUnit 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit:"
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
         Left            =   180
         TabIndex        =   23
         Top             =   1560
         Width           =   315
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Double Click an Item to Edit."
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
         Height          =   210
         Left            =   3870
         TabIndex        =   21
         Top             =   4920
         Width           =   2280
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "LIS Result Code:"
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
         TabIndex        =   18
         Top             =   2730
         Width           =   1185
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Active:"
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
         Left            =   210
         TabIndex        =   17
         Top             =   3120
         Width           =   510
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "LIS Exam Code:"
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
         TabIndex        =   16
         Top             =   2310
         Width           =   1125
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Display Name:"
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
         TabIndex        =   15
         Top             =   1950
         Width           =   1020
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument Exam Name:"
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
         TabIndex        =   14
         Top             =   1140
         Width           =   1680
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument Exam Code:"
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
         Left            =   180
         TabIndex        =   13
         Top             =   750
         Width           =   1650
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument ID:"
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
         Left            =   180
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmLabExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mConnection As New ADODB.Connection
Private rstExams As ADODB.Recordset
Private AddMode As Boolean

Private Sub cboActive_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmdAddSave_Click()
Dim rstCheck As ADODB.Recordset
Dim rstSaveData As ADODB.Recordset
If AddMode Then
    If Len(txtResultCode.Text) = 0 Or Len(txtInstrumentExamName.Text) = 0 Or Len(txtLISExamCode.Text) = 0 Then
        MsgBox "Unable to save data.", vbOKOnly + vbInformation, "Data Check"
        txtInstrumentExamName.SetFocus
    Else
        Set rstCheck = mConnection.Execute("qryCheckExam '" & txtInstrumentID.Text & "','" & txtInstrumentExamName.Text & "','" & txtLISExamCode.Text & "','" & txtResultCode.Text & "'")
        If rstCheck.RecordCount = 0 Then
            Set rstSaveData = mConnection.Execute("qryInsertExam '" & txtInstrumentID.Text & "','" & txtRefNum.Text & "','" & txtInstrumentExamName.Text & "','" & txtLISExamCode.Text & "','" & txtDisplayName.Text & "'," & cboActive.Text & ",'" & txtResultCode.Text & "','" & txtUnit.Text & "'")
            clearValues
            txtInstrumentExamName.SetFocus
            msfgExams.Refresh
        Else
            MsgBox "Record Already Exists.", vbOKOnly + vbInformation, "Record Check"
        End If
    End If
Else
    Set rstCheck = mConnection.Execute("qryCheckExam2 '" & txtInstrumentID.Text & "','" & txtRefNum.Text & "','" & txtLISExamCode.Text & "','" & txtResultCode.Text & "'")
    If rstCheck.RecordCount = 0 Then
        Set rstSaveData = mConnection.Execute("qryUpdateExam '" & txtInstrumentID.Text & "','" & txtRefNum.Text & "','" & txtInstrumentExamName.Text & "','" & txtLISExamCode.Text & "','" & txtDisplayName.Text & "'," & cboActive.Text & ",'" & txtResultCode.Text & "','" & txtUnit.Text & "'")
        cmdCancel.Visible = False
        cmdAddSave.Caption = "Add"
        AddMode = True
        msfgExams.Enabled = True
        clearValues
        cmdClose.Cancel = True
        lblStatus.Caption = "Double Click an Item to Edit."
    Else
        MsgBox "Record Already Exists.", vbOKOnly + vbInformation, "Record Check"
    End If
End If
setValues
End Sub

Private Sub cmdCancel_Click()
cmdCancel.Visible = False
cmdAddSave.Caption = "Add"
AddMode = True
msfgExams.Enabled = True
clearValues
cmdClose.Cancel = True
lblStatus.Caption = "Double Click an Item to Edit."
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub OpenLocalConnection()
On Error GoTo errFound
    Screen.ActiveForm.MousePointer = 11
    mConnection.ConnectionString = getConStr
    mConnection.ConnectionTimeout = 30
    mConnection.CursorLocation = adUseClient
    mConnection.Open
errFound:
    If Err.Number <> 0 Then
        Select Case Err.Number
            Case -2147467259
                MsgBox "Unable to connect to server" & vbCrLf & vbCrLf & "Please check database settings.", vbExclamation + vbOKOnly, "Database Error"
            Case Else
                MsgBox Err.Description & vbCrLf & "Please Contact your Systems Administrator.", vbExclamation + vbOKOnly, "Database Error"
        End Select
    End If
    Screen.ActiveForm.MousePointer = 0
End Sub

Private Function getConStr() As String
Dim tempStr As String
Dim mUserName As String
Dim mDatabase As String
Dim mPassword  As String
    tempStr = GetSetting("lab-i", "LOCAL", "UserName", "")
    If Len(Trim(tempStr)) > 0 Then
        mUserName = Blowfish.DecryptString(tempStr, myKey, True)
    Else
        mUserName = ""
    End If
    tempStr = GetSetting("lab-i", "LOCAL", "DATABASEPASSWORD", "")
    If Len(Trim(tempStr)) > 0 Then
        mPassword = Blowfish.DecryptString(tempStr, myKey, True)
    Else
        mPassword = ""
    End If
    tempStr = GetSetting("lab-i", "LOCAL", "DATABASE", "")
    If Len(Trim(tempStr)) > 0 Then
        mDatabase = Blowfish.DecryptString(tempStr, myKey, True)
    Else
        mDatabase = ""
    End If
    getConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mDatabase & ";User Id=" & mUserName & ";Password=;" & "Jet OLEDB:Database Password=" & mPassword & ";"
End Function

Private Sub Form_Load()
    OpenLocalConnection
    setValues
    AddMode = True
End Sub

Private Function getRefNum() As String
Dim rstData As ADODB.Recordset
    Set rstData = mConnection.Execute("Select max(examrefnum) as refnum from tbCoLabExam")
    If rstData.State = 1 Then
        If rstData.RecordCount > 0 Then
            getRefNum = IIf(IsNull(rstData!refnum), "0", rstData!refnum)
            getRefNum = Val(getRefNum) + 1
        Else
            getRefNum = "1"
        End If
    End If
    
End Function

Private Sub setValues()
    txtRefNum.Text = getRefNum
    cboActive.ListIndex = 0
    Set rstExams = mConnection.Execute("qrySelectExams '" & txtInstrumentID.Text & "'")
    If rstExams.RecordCount > 0 Then
        setGrid False
        Set msfgExams.Recordset = rstExams
    Else
        setGrid
    End If
End Sub

Private Sub setItems(ItemId As String)
Dim rstItems As New ADODB.Recordset
Set rstItems = mConnection.Execute("qryGetExamItem " & ItemId)
With rstItems
    txtRefNum.Text = !ExamRefNum
    txtInstrumentExamName.Text = !InstExamName & ""
    txtDisplayName.Text = !DisplayName & ""
    txtLISExamCode.Text = !HisExamCode & ""
    txtResultCode.Text = !HisResultCode & ""
    txtUnit.Text = !Unit & ""
    If !Active Then cboActive.ListIndex = 0 Else cboActive.ListIndex = 1
End With
End Sub

Private Sub clearValues()
    txtRefNum.Text = getRefNum
    txtInstrumentExamName.Text = ""
    txtDisplayName.Text = ""
    txtLISExamCode.Text = ""
    txtResultCode.Text = ""
    txtUnit.Text = ""
    cboActive.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
mConnection.Close
End Sub

Private Sub msfgExams_DblClick()
With msfgExams
    setItems .TextMatrix(.Row, 0)
    AddMode = False
    cmdAddSave.Caption = "Update"
    cmdCancel.Visible = True
    msfgExams.Enabled = False
    cmdCancel.Cancel = True
    txtInstrumentExamName.SetFocus
    lblStatus.Caption = "Editing Instrument Exam Number: " & .TextMatrix(.Row, 0)
End With
End Sub

Private Sub txtDisplayName_GotFocus()
highLight
End Sub

Private Sub txtDisplayName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtInstrumentExamName_GotFocus()
highLight
End Sub

Private Sub txtInstrumentExamName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtInstrumentExamName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtLISExamCode_GotFocus()
highLight
End Sub

Private Sub txtLISExamCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub setGrid(Optional isclear As Boolean = True)
With msfgExams
If isclear Then
    .Clear
    .Cols = 6
    .Rows = 2
    .TextMatrix(0, 0) = "Code"
    .TextMatrix(0, 1) = "Exam"
    .TextMatrix(0, 2) = "Instrument Exam"
    .TextMatrix(0, 3) = "Exam Code"
    .TextMatrix(0, 4) = "Result Code"
    .TextMatrix(0, 5) = "Active"
Else

End If
    .ColWidth(0) = 500
    .ColWidth(1) = 2400
    .ColWidth(2) = 1300
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 800
End With
End Sub

Private Sub txtResultCode_GotFocus()
    highLight
End Sub

Private Sub txtResultCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtUnit_GotFocus()
    highLight
End Sub
