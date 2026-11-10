VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLocAvailableRates 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Available Exam Rates"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraBody 
      Caption         =   "Frame1"
      Height          =   405
      Index           =   0
      Left            =   1365
      TabIndex        =   9
      Top             =   3990
      Visible         =   0   'False
      Width           =   945
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgLocRates 
      Height          =   3375
      Left            =   90
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   420
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   5953
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483634
      GridColorFixed  =   -2147483634
      GridColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      FillStyle       =   1
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdSingleRight 
      Caption         =   ">"
      Height          =   465
      Left            =   3645
      TabIndex        =   1
      Top             =   1380
      Width           =   555
   End
   Begin VB.CommandButton cmdAllLeft 
      Caption         =   "<<"
      Height          =   465
      Left            =   3645
      TabIndex        =   2
      Top             =   2370
      Width           =   555
   End
   Begin VB.CommandButton cmdAllRight 
      Caption         =   ">>"
      Height          =   465
      Left            =   3645
      TabIndex        =   4
      Top             =   3045
      Width           =   555
   End
   Begin VB.CommandButton cmdSingleLeft 
      Caption         =   "<"
      Height          =   465
      Left            =   3645
      TabIndex        =   0
      Top             =   690
      Width           =   555
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   555
      Left            =   6705
      TabIndex        =   6
      Top             =   3930
      Width           =   1080
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgAllExams 
      Height          =   3375
      Left            =   4380
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   420
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   5953
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483634
      GridColorFixed  =   -2147483634
      GridColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      FillStyle       =   1
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lbltext 
      BackStyle       =   0  'Transparent
      Caption         =   "ALL EXAM RATES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   4380
      TabIndex        =   5
      Top             =   180
      Width           =   4125
   End
   Begin VB.Label lbltext 
      BackStyle       =   0  'Transparent
      Caption         =   "SELECTED EXAM RATES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   180
      Width           =   4125
   End
End
Attribute VB_Name = "frmLocAvailableRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PreRates As String
Private SelRates As String
Private ctrMsfgRow As Integer
    
Public Property Let InitialRateID(ByVal vNewValue As String)
    PreRates = vNewValue
End Property

Public Property Get SelectedRates() As String
    SelectedRates = SelRates
End Property

Private Sub cmdAllLeft_Click()
    With msfgLocRates
        For ctrMsfgRow = 0 To msfgAllExams.Rows - 1
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .Col = 0
            .CellAlignment = 0
            .TextMatrix(.Row, 0) = msfgAllExams.TextMatrix(ctrMsfgRow, 0)
            .TextMatrix(.Row, 1) = msfgAllExams.TextMatrix(ctrMsfgRow, 1)
        Next
        msfgAllExams.Rows = 0
        cmdSingleLeft.Enabled = False
        cmdAllLeft.Enabled = False
        cmdSingleRight.Enabled = True
        cmdAllRight.Enabled = True
        .Row = 0
        .ColSel = 1
    End With
    
End Sub

Private Sub cmdAllRight_Click()
    With msfgAllExams
        For ctrMsfgRow = 0 To msfgLocRates.Rows - 1
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .Col = 0
            .CellAlignment = 0
            .TextMatrix(.Row, 0) = msfgLocRates.TextMatrix(ctrMsfgRow, 0)
            .TextMatrix(.Row, 1) = msfgLocRates.TextMatrix(ctrMsfgRow, 1)
        Next
        msfgLocRates.Rows = 0
        cmdSingleLeft.Enabled = True
        cmdAllLeft.Enabled = True
        cmdSingleRight.Enabled = False
        cmdAllRight.Enabled = False
        .Row = 0
        .ColSel = 1
    End With
End Sub

Private Sub cmdDone_Click()
    SelRates = ""
    With msfgLocRates
        If .Rows > 0 Then
            For ctrMsfgRow = 0 To .Rows - 1
                SelRates = IIf(Len(SelRates) > 0, _
                            SelRates & "," & .TextMatrix(ctrMsfgRow, 0), .TextMatrix(ctrMsfgRow, 0))
            Next
        End If
    End With
    Unload Me
End Sub

Private Sub cmdSingleLeft_Click()
    With msfgLocRates
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .Col = 0
        .CellAlignment = 0
        .TextMatrix(.Row, 0) = msfgAllExams.TextMatrix(msfgAllExams.Row, 0)
        .TextMatrix(.Row, 1) = msfgAllExams.TextMatrix(msfgAllExams.Row, 1)
        On Error Resume Next
        If msfgAllExams.Rows > 1 Then
            msfgAllExams.RemoveItem (msfgAllExams.Row)
            msfgAllExams.ColSel = 1
            'msfgAllExams.SetFocus
        ElseIf msfgAllExams.Rows = 1 Then
            msfgAllExams.Rows = 0
            cmdSingleLeft.Enabled = False
            cmdAllLeft.Enabled = False
        End If
        cmdSingleRight.Enabled = True
        cmdAllRight.Enabled = True
        .ColSel = 1
    End With
End Sub

Private Sub cmdSingleRight_Click()
    With msfgAllExams
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .Col = 0
        .CellAlignment = 0
        .TextMatrix(.Row, 0) = msfgLocRates.TextMatrix(msfgLocRates.Row, 0)
        .TextMatrix(.Row, 1) = msfgLocRates.TextMatrix(msfgLocRates.Row, 1)
        If msfgLocRates.Rows > 1 Then
            msfgLocRates.RemoveItem (msfgLocRates.Row)
            'msfgLocRates.SetFocus
            msfgLocRates.ColSel = 1
        ElseIf msfgLocRates.Rows = 1 Then
            msfgLocRates.Rows = 0
            cmdSingleRight.Enabled = False
            cmdAllRight.Enabled = False
        End If
        cmdSingleLeft.Enabled = True
        cmdAllLeft.Enabled = True
        .ColSel = 1
    End With
End Sub

Private Sub Form_Load()

    setFormColors Me
    With Me
        .BackColor = GetSetting("LMS", "FormColors", "BodyFrame", &H6B4B01)
        .ForeColor = GetSetting("LMS", "FormColors", "TitleText", &HFFFFFF)
        .BackColor = GetSetting("LMS", "FormColors", "BodyFrame", &H6B4B01)
        .ForeColor = GetSetting("LMS", "FormColors", "TitleText", &HFFFFFF)
        .BackColor = GetSetting("LMS", "FormColors", "BodyFrame", &H6B4B01)
        .ForeColor = GetSetting("LMS", "FormColors", "TitleText", &HFFFFFF)
    End With
    
    SetAllExamLst
    SetLocRates
    
    If msfgLocRates.Rows = 0 Then
        cmdSingleRight.Enabled = False
        cmdAllRight.Enabled = False
    End If
    
    If msfgAllExams.Rows = 0 Then
        cmdSingleLeft.Enabled = False
        cmdAllLeft.Enabled = False
    End If
    
End Sub

Private Sub SetAllExamLst()
    
    Dim rsAllExam As ADODB.Recordset
    Dim ctrExamRow As Integer
    
    Set rsAllExam = myConnection.Execute("SELECT * FROM tbLabExamRatesClass")

    msfgAllExams.ColWidth(0) = 400
    msfgAllExams.ColWidth(1) = 3000

    If rsAllExam.RecordCount > 0 Then
        With rsAllExam
            For ctrExamRow = 0 To .RecordCount - 1
                msfgAllExams.Rows = msfgAllExams.Rows + 1
                msfgAllExams.Row = msfgAllExams.Rows - 1
                msfgAllExams.Col = 0
                msfgAllExams.CellAlignment = 0
                msfgAllExams.TextMatrix(msfgAllExams.Row, 0) = .Fields("ExamClassID")
                msfgAllExams.TextMatrix(msfgAllExams.Row, 1) = .Fields("ExamClassName")
                .MoveNext
            Next
        End With
        msfgAllExams.Row = 0
    End If
End Sub

Private Sub SetLocRates()
    Dim arryPreRates() As String
    Dim ctrArryPreRates As Integer
    
    msfgLocRates.ColWidth(0) = 400
    msfgLocRates.ColWidth(1) = 3000
    
    If Len(PreRates) > 0 Then
        arryPreRates = Split(PreRates, ",")
        If msfgAllExams.Rows > 0 Then
            With msfgLocRates
                For ctrArryPreRates = 0 To UBound(arryPreRates)
                    If Len(Trim(arryPreRates(ctrArryPreRates))) > 0 Then
                        For ctrMsfgRow = 0 To msfgAllExams.Rows - 1
                            If arryPreRates(ctrArryPreRates) = msfgAllExams.TextMatrix(ctrMsfgRow, 0) Then
                                msfgAllExams.Row = ctrMsfgRow
                                cmdSingleLeft_Click
                                Exit For
                            End If
                        Next
                    End If
                Next
            End With
        End If
    End If
    
End Sub

