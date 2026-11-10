VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmInstrumentResult 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Instrument Result"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   14835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Height          =   405
      Left            =   11910
      TabIndex        =   4
      Top             =   8070
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   13350
      TabIndex        =   3
      Top             =   8070
      Width           =   1485
   End
   Begin VB.CheckBox chkExam 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Top             =   630
      Visible         =   0   'False
      Width           =   195
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
      Height          =   7965
      Left            =   10020
      TabIndex        =   0
      Top             =   30
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   14049
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   12615680
      ForeColorFixed  =   16777215
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgInstResult 
      Height          =   7995
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   14102
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      BackColorFixed  =   12615680
      ForeColorFixed  =   16777215
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   9570
      Picture         =   "frmInstrumentResult.frx":0000
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   9570
      Picture         =   "frmInstrumentResult.frx":0442
      Top             =   5280
      Width           =   480
   End
End
Attribute VB_Name = "frmInstrumentResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ySampleNumber As String
Dim ySpecimen As String
Dim yRequestid As String
Dim yUsedResult As String
Dim yShowForm As Boolean
Dim yInstrumentID As String
Dim startHighlight  As Boolean
Dim myTop() As String
Dim multipleCheck As Boolean
Dim HasValue As Boolean

Private Sub chkExam_Click(Index As Integer)
With msfgInstResult
    .Redraw = False
    If chkExam(Index).Value = 1 Then
        If HasSameExam(Index) Then
            valInstExamChk Index
        End If
        If setInstResult(1, .TextMatrix(Index, 5), .TextMatrix(Index, 6), .TextMatrix(Index, 3), .TextMatrix(Index, 4), .TextMatrix(Index, 5), .TextMatrix(Index, 10)) Then
            .TextMatrix(Index, 0) = "                              1"
        Else
            chkExam(Index).Value = 0
        End If
    Else
        setInstResult 0, .TextMatrix(Index, 5), .TextMatrix(Index, 6), "", "", "", .TextMatrix(Index, 10)
        .TextMatrix(Index, 0) = "                              0"
    End If
    .row = Index
    .ColSel = 9
    .Redraw = True
End With
End Sub

Private Sub highlightChecked()
Dim ctr As Integer
With msfgInstResult
    For ctr = 1 To .Rows - 1
        .row = ctr
        .Col = 0
        If Val(.TextMatrix(ctr, 0)) = 1 Then
            .CellBackColor = vbHighlight
            .CellForeColor = vbHighlightText
            .CellForeColor = vbHighlight
            .Col = 1
            .CellBackColor = vbHighlight
            .CellForeColor = vbHighlightText
            .Col = 2
            .CellBackColor = vbHighlight
            .CellForeColor = vbHighlightText
            .Col = 5
            .CellBackColor = vbHighlight
            .CellForeColor = vbHighlightText
            .Col = 6
            .CellBackColor = vbHighlight
            .CellForeColor = vbHighlightText
        Else
            .CellBackColor = vbWindowBackground
            .CellForeColor = vbWindowBackground
            .Col = 1
            .CellBackColor = vbWindowBackground
            .CellForeColor = vbWindowText
            .Col = 2
            .CellBackColor = vbWindowBackground
            .CellForeColor = vbWindowText
            .Col = 5
            .CellBackColor = vbWindowBackground
            .CellForeColor = vbWindowText
            .Col = 6
            .CellBackColor = vbWindowBackground
            .CellForeColor = vbWindowText
        End If
    Next ctr
End With
End Sub

Private Sub cmdCancel_Click()
    ReturnValueStr = ""
    Unload Me
End Sub

Private Sub cmdOk_Click()
Dim ctr As Integer
Dim myStr As String
If HasValue Then
    myStr = ""
    With msfgResult
        For ctr = 1 To .Rows - 1
            If Len(.TextMatrix(ctr, 2)) > 0 Then
              myStr = myStr & .TextMatrix(ctr, 9) & "|" & .TextMatrix(ctr, 2) & "|" & .TextMatrix(ctr, 13) & "|" & .TextMatrix(ctr, 14) & "|" & .TextMatrix(ctr, 15) & "|" & .TextMatrix(ctr, 25) & vbCrLf
            End If
        Next ctr
    End With
    If Len(myStr) > 0 Then
        ReturnValueStr = myStr
    Else
        ReturnValueStr = ""
    End If
End If
Unload Me
End Sub

Private Sub Form_Load()
ReturnValueStr = ""
HasValue = False
startHighlight = False
LoadInstrumentResults
startHighlight = True
multipleCheck = False
If Not yShowForm Then cmdOk_Click
'displaycols
End Sub

Private Sub LoadInstrumentResults()
Dim rstTemp As adodb.Recordset
Dim ctr As Integer
Dim ctr2 As Integer
Dim tempStr1, tempStr2 As String
Dim Used As Boolean
Dim arrStr() As String
Dim arrStr2() As String
Dim tempStrRow As String
tempStr1 = ""
tempStr2 = ""
tempStrRow = ""
Set rstTemp = myConnection.Execute("sp_DisplayResultEntry2 '" & yRequestid & "','" & yInstrumentID & "','" & ySpecimen & "'")
If rstTemp.RecordCount > 0 Then
    HasValue = True
    msfgResult.Visible = True
    Set msfgResult.DataSource = rstTemp
    
    Set rstTemp = myConnection.Execute("sp_InstrumentResults_Select '" & ySampleNumber & "'")
    Set msfgInstResult.DataSource = rstTemp
    With msfgInstResult
        If rstTemp.RecordCount > 0 Then
            .Cols = .Cols + 1
            For ctr = 1 To .Rows - 1
                Load chkExam(ctr)
                chkExam(ctr).Enabled = True
                chkExam(ctr).Visible = True
                chkExam(ctr).ZOrder 0
                .Col = 0
                'chkExam(ctr).Move .Left + .CellLeft + (.CellWidth / 2) - (chkExam(ctr).Width / 2), .Top + .CellTop + (.CellHeight / 2) - (chkExam(ctr).Height / 2) + (.RowHeight(ctr) * (ctr - 1))
                chkExam(ctr).Move .CellLeft + (chkExam(ctr).Width / 2), .Top + .CellTop + (.CellHeight / 2) - (chkExam(ctr).Height / 2) + (.RowHeight(ctr) * (ctr - 1))
                chkExam(ctr).Visible = .RowIsVisible(ctr)
                tempStrRow = tempStrRow & .Top + .CellTop + (.CellHeight / 2) - (chkExam(ctr).Height / 2) + (.RowHeight(ctr) * (ctr - 1)) & "|"
            Next ctr
            If Len(Trim(tempStrRow)) > 0 Then
                tempStrRow = Left(tempStrRow, Len(tempStrRow) - 1)
                myTop = Split(tempStrRow, "|")
            End If
        End If
        For ctr = 1 To .Rows - 1
            If Len(Trim(tempStr1)) > 0 Then
                If tempStr1 = .TextMatrix(ctr, 3) And tempStr2 = .TextMatrix(ctr, 4) Then
                    .TextMatrix(ctr, 1) = ""
                    .TextMatrix(ctr, 2) = ""
                Else
                    tempStr1 = .TextMatrix(ctr, 3)
                    tempStr2 = .TextMatrix(ctr, 4)
                    .TextMatrix(ctr, 1) = Format(tempStr1, "dd-MMM-yyyy HH:mm")
                    .TextMatrix(ctr, 2) = tempStr2
                End If
            Else
                tempStr1 = .TextMatrix(ctr, 3)
                tempStr2 = .TextMatrix(ctr, 4)
                .TextMatrix(ctr, 1) = Format(tempStr1, "dd-MMM-yyyy HH:mm")
                .TextMatrix(ctr, 2) = tempStr2
            End If
            .TextMatrix(ctr, 9) = getExamName(.TextMatrix(ctr, 5), .TextMatrix(ctr, 10))
'            If Len(Trim(.TextMatrix(ctr, 9))) = 0 Then
'                .RowHeight(ctr) = 0
'                chkExam(ctr).Visible = False
'            End If
        Next ctr
        Used = hasUsed
        tempStr1 = ""
        tempStr2 = ""
        For ctr = 1 To .Rows - 1
            If Not Used Then
                If Len(Trim(yUsedResult)) > 0 And InStr(1, yUsedResult, "|") > 0 Then
                    arrStr = Split(yUsedResult, vbCrLf)
                    For ctr2 = 0 To UBound(arrStr) - 1
                        arrStr2 = Split(arrStr(ctr2), "|")
                        If arrStr2(2) = .TextMatrix(ctr, 3) And arrStr2(3) = .TextMatrix(ctr, 4) And arrStr2(4) = .TextMatrix(ctr, 5) Then
                            chkExam(ctr).Value = 1
                        Else
                            .Col = 0
                            .row = ctr
                            .CellForeColor = .CellBackColor
                        End If
                    Next ctr2
                Else
                    If Len(Trim(tempStr1)) > 0 Then
                        If tempStr1 = .TextMatrix(ctr, 3) And tempStr2 = .TextMatrix(ctr, 4) Then
                            chkExam(ctr).Value = 1
                        Else
                            Debug.Print ctr & "-" & .TextMatrix(ctr, 9)
                            If Not HasSameExam(ctr) Then
                                chkExam(ctr).Value = 1
                            End If
                            .Col = 0
                            .row = ctr
                            .CellForeColor = .CellBackColor
        '                    Exit For
                        End If
                    Else
                        tempStr1 = .TextMatrix(ctr, 3)
                        tempStr2 = .TextMatrix(ctr, 4)
                        chkExam(ctr).Value = 1
                    End If
                End If
            Else
                If Len(Trim(yUsedResult)) > 0 And InStr(1, yUsedResult, "|") > 0 Then
                    arrStr = Split(yUsedResult, vbCrLf)
                    For ctr2 = 0 To UBound(arrStr) - 1
                        arrStr2 = Split(arrStr(ctr2), "|")
                        If arrStr2(2) = .TextMatrix(ctr, 3) And arrStr2(3) = .TextMatrix(ctr, 4) And arrStr2(4) = .TextMatrix(ctr, 5) Then
                            chkExam(ctr).Value = 1
                        Else
                            .Col = 0
                            .row = ctr
                            .CellForeColor = .CellBackColor
                        End If
                    Next ctr2
                Else
                    If Trim(Format(.TextMatrix(ctr, 7), ">")) = "TRUE" Then
                        chkExam(ctr).Value = 1
                    Else
                        chkExam(ctr).Value = 0
                    End If
                End If
                .Col = 0
                .row = ctr
                .CellForeColor = .CellBackColor
            End If
        Next ctr
    End With
    setGrids
    startHighlight = True
Else
    HasValue = False
    yShowForm = False
End If
End Sub

Private Function hasUsed() As Boolean
Dim ctr As Integer
Dim tempBool As Boolean
tempBool = False
With msfgInstResult
    For ctr = 1 To .Rows - 1
        If Format(.TextMatrix(ctr, 7), ">") = "TRUE" Then
            tempBool = True
            Exit For
        End If
    Next ctr
End With
hasUsed = tempBool
End Function

Public Property Let wSampleNumber(ByVal vNewValue As String)
    ySampleNumber = vNewValue
End Property

Public Property Let wSpecimen(ByVal vNewValue As String)
    ySpecimen = vNewValue
End Property

Public Property Let wRequestID(ByVal vNewValue As String)
    yRequestid = vNewValue
End Property

Private Sub setGrids()
With msfgInstResult
    .ColWidth(0) = 400
    .ColWidth(1) = 2000
    .ColWidth(2) = 1600
    .ColWidth(3) = 0
    .ColWidth(4) = 0
    .ColWidth(5) = 1500
    .ColWidth(6) = 1400
    .ColWidth(7) = 0
    .ColWidth(8) = 0
    .ColWidth(9) = 2100
    .ColWidth(10) = 0
    .ColWidth(11) = 0
    .ColWidth(12) = 0
    If .Rows > 1 Then .Height = .RowHeight(1) * 29
End With
With msfgResult
    .Height = msfgInstResult.Height
    .ColWidth(0) = 0
    .ColWidth(1) = 2500
    .ColWidth(2) = 2000
    .ColWidth(3) = 0
    .ColWidth(4) = 0
    .ColWidth(5) = 0
    .ColWidth(6) = 0
    .ColWidth(7) = 0
    .ColWidth(8) = 0
    .ColWidth(9) = 0
    .ColWidth(10) = 0
    .ColWidth(11) = 0
    .ColWidth(12) = 0
    .ColWidth(13) = 0
    .ColWidth(14) = 0
    .ColWidth(15) = 0
    .ColWidth(16) = 0
    .ColWidth(17) = 0
    .ColWidth(18) = 0
    .ColWidth(19) = 0
    .ColWidth(20) = 0
    .ColWidth(21) = 0
    .ColWidth(22) = 0
    .ColWidth(23) = 0
    .ColWidth(24) = 0
    .ColWidth(25) = 0
    .ColWidth(26) = 0
    .ColWidth(27) = 0
    .ColWidth(28) = 0
End With
End Sub

Private Sub displaycols()
Dim ctr As Integer
With msfgResult
    For ctr = 0 To .Cols - 1
        Debug.Print ctr & " - " & .TextMatrix(0, ctr)
    Next ctr
End With
End Sub

Private Function setInstResult(mTrans As Integer, mExamName As String, mResult As String, mDate As String, mHostID As String, mInstExam As String, mPeriodRun As String) As Boolean
Dim ctr As Integer
Dim hasInstResult As Boolean
mExamName = Format(mExamName, ">")
hasInstResult = False
With msfgResult
    For ctr = 1 To .Rows - 1
        If InStr(1, .TextMatrix(ctr, 6), "<" & mExamName & ">") > 0 And .TextMatrix(ctr, 25) = mPeriodRun Then
            If mTrans = 0 Then
                .TextMatrix(ctr, 2) = ""
                .TextMatrix(ctr, 13) = ""
                .TextMatrix(ctr, 14) = ""
                .TextMatrix(ctr, 15) = ""
            Else
                .TextMatrix(ctr, 2) = mResult
                .TextMatrix(ctr, 13) = mDate
                .TextMatrix(ctr, 14) = mHostID
                .TextMatrix(ctr, 15) = mInstExam
                hasInstResult = True
            End If
        End If
    Next ctr
End With
setInstResult = hasInstResult
End Function

Private Function HasSameExam(mIndex As Integer) As Boolean
Dim ctr As Integer
Dim newDate As String
Dim newTest As String
Dim newHostCode As String
Dim newPeriodRun As String
Dim SameExam As Boolean
SameExam = False
With msfgInstResult
    newDate = .TextMatrix(mIndex, 3)
    newTest = .TextMatrix(mIndex, 9)
    newPeriodRun = .TextMatrix(mIndex, 10)
    newHostCode = .TextMatrix(mIndex, 4)
    For ctr = 1 To .Rows - 1
        If ctr <> mIndex And Val(.TextMatrix(ctr, 0)) = 1 Then
            If .TextMatrix(ctr, 3) <> newDate And .TextMatrix(ctr, 9) = newTest And .TextMatrix(ctr, 10) = newPeriodRun Then
'            If .TextMatrix(ctr, 5) = newTest And .TextMatrix(ctr, 10) = newPeriodRun Then
                SameExam = True
                Exit For
            ElseIf .TextMatrix(ctr, 4) <> newHostCode And .TextMatrix(ctr, 3) = newDate And .TextMatrix(ctr, 9) = newTest And .TextMatrix(ctr, 10) = newPeriodRun Then
'            ElseIf .TextMatrix(ctr, 4) <> newHostCode And .TextMatrix(ctr, 5) = newTest And .TextMatrix(ctr, 10) = newPeriodRun Then
                SameExam = True
                Exit For
            End If
        End If
    Next ctr
End With
HasSameExam = SameExam
End Function

Private Sub valInstExamChk(mIndex As Integer)
Dim ctr As Integer
Dim newDate As String
Dim newTest As String
Dim newHostCode As String
Dim newPeriodRun As String
With msfgInstResult
    newDate = .TextMatrix(mIndex, 3)
    newTest = .TextMatrix(mIndex, 9)
    newHostCode = .TextMatrix(mIndex, 4)
    newPeriodRun = .TextMatrix(mIndex, 10)
    For ctr = 1 To .Rows - 1
        If ctr <> mIndex And Val(.TextMatrix(ctr, 0)) = 1 Then
            If .TextMatrix(ctr, 3) <> newDate And .TextMatrix(ctr, 9) = newTest And .TextMatrix(ctr, 10) = newPeriodRun Then
'            If .TextMatrix(ctr, 5) = newTest And .TextMatrix(ctr, 10) = newPeriodRun Then
                setCheckValues newDate, newHostCode
                Exit For
            ElseIf .TextMatrix(ctr, 4) <> newHostCode And .TextMatrix(ctr, 3) = newDate And .TextMatrix(ctr, 5) = newTest And .TextMatrix(ctr, 10) = newPeriodRun Then
'            ElseIf .TextMatrix(ctr, 4) <> newHostCode And .TextMatrix(ctr, 5) = newTest And .TextMatrix(ctr, 10) = newPeriodRun Then
                setCheckValues newDate, newHostCode
                Exit For
            End If
        End If
    Next ctr
End With
End Sub

Private Sub setCheckValues(mDate As String, mHostCode As String)
Dim ctr As Integer
With msfgInstResult
    For ctr = 1 To .Rows - 1
        If .TextMatrix(ctr, 3) = mDate And .TextMatrix(ctr, 4) = mHostCode Then
            chkExam(ctr).Value = 1
        Else
            chkExam(ctr).Value = 0
        End If
    Next ctr
End With
End Sub

Private Sub msfgInstResult_Click()
    With msfgInstResult
        If Trim(.TextMatrix(.row, 0)) = "1" Then
            chkExam(.row).Value = 0
        Else
            chkExam(.row).Value = 1
        End If
    End With
End Sub

Private Sub msfgInstResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
    With msfgInstResult
        If Trim(.TextMatrix(.row, 0)) = "1" Then
            chkExam(.row).Value = 0
        Else
            chkExam(.row).Value = 1
        End If
    End With
End If
End Sub

Public Property Let wUsedResult(ByVal vNewValue As String)
    yUsedResult = vNewValue
End Property

Public Property Let wShowForm(ByVal vNewValue As Boolean)
    yShowForm = vNewValue
End Property

Public Property Let wInstrumentID(ByVal vNewValue As String)
    yInstrumentID = vNewValue
End Property

Private Sub msfgInstResult_Scroll()
Dim ctr As Integer
Dim ctr2 As Integer
With msfgInstResult
    For ctr = 1 To .Rows - 1
        chkExam(ctr).Visible = .RowIsVisible(ctr)
    Next ctr
    ctr2 = 0
    For ctr = .TopRow To .Rows - 1
        chkExam(ctr).Top = myTop(ctr2)
        ctr2 = ctr2 + 1
    Next ctr
End With
End Sub

Private Function getExamName(mExamCode As String, mPeriodRun As String) As String
Dim ctr As Integer
With msfgResult
    If .Rows > 0 Then
        For ctr = 1 To .Rows - 1
            If InStr(1, Format(.TextMatrix(ctr, 6), ">"), "<" & Format(mExamCode, ">") & ">") > 0 And .TextMatrix(ctr, 25) = mPeriodRun Then
                getExamName = .TextMatrix(ctr, 1)
                Exit For
            End If
        Next ctr
    End If
End With
End Function
