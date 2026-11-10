VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDelReq 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Delete Request"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   495
      Left            =   780
      TabIndex        =   11
      Top             =   6195
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   7665
      TabIndex        =   10
      Top             =   4230
      Width           =   1215
   End
   Begin VB.TextBox txtmsfgExam 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   525
      Left            =   2880
      MaxLength       =   30
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   8190
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   4995
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   -90
      Width           =   9165
      Begin VB.CommandButton cmdRefreshList 
         Caption         =   "Refresh List"
         Height          =   345
         Left            =   4950
         TabIndex        =   2
         Top             =   315
         Width           =   1260
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   345
         Left            =   1185
         TabIndex        =   0
         Top             =   300
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   112394243
         CurrentDate     =   41283
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   345
         Left            =   3465
         TabIndex        =   1
         Top             =   300
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   112394243
         CurrentDate     =   41283
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   3435
         Left            =   255
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   810
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   6059
         _Version        =   393216
         Cols            =   3
         FocusRect       =   0
         HighLight       =   2
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
         _Band(0).Cols   =   3
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "End Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   2610
         TabIndex        =   8
         Top             =   345
         Width           =   1215
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   345
         Width           =   1215
      End
   End
   Begin VB.Frame fraTitle 
      Caption         =   "Frame1"
      Height          =   495
      Left            =   645
      TabIndex        =   3
      Top             =   7035
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imglstIcon 
      Left            =   4770
      Top             =   8295
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReqExam.frx":0000
            Key             =   "newrequest"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReqExam.frx":03FA
            Key             =   "checkin"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReqExam.frx":081E
            Key             =   "ongoing"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReqExam.frx":0C43
            Key             =   "ready"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReqExam.frx":1047
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReqExam.frx":144F
            Key             =   "new"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReqExam.frx":184B
            Key             =   "printed"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReqExam.frx":1C6A
            Key             =   "saved"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReqExam.frx":207D
            Key             =   "unapprove"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgExamDet 
      Height          =   2445
      Left            =   4485
      TabIndex        =   5
      Top             =   4980
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   4313
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
End
Attribute VB_Name = "frmDelReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intRowCtr As Integer
Dim intColCtr As Integer
Dim rsTest As New ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    setFormColors Me
    dtpStartDate.Value = Now
    dtpEndDate.Value = Now
    
End Sub

Private Sub setmsfgExam(strRequestID As String)

    With msfgResult
        strSQLStatement = "EXEC sp_tbLabRequestValues_select '" & strRequestID & "'"
        Set rsTest = myConnection.Execute(strSQLStatement)

        Set .DataSource = rsTest
        .Refresh
        .Rows = .Rows + 1
        .Enabled = True
        
        For intRowCtr = 0 To .Cols - 1
            .ColWidth(intRowCtr) = 0
        Next

        .TextArray(0) = "Status"
        
        For intRowCtr = 1 To .Rows - 1
            .Row = intRowCtr
            .Col = 0
            .CellAlignment = 0
            .CellBackColor = vbWhite
            .TextMatrix(intRowCtr, 0) = .TextMatrix(intRowCtr, 6)
            
            Select Case .TextMatrix(.Row, .Col)
                Case ""
                    Set .CellPicture = imglstIcon.ListImages(6).Picture
                Case "0"
                    Set .CellPicture = imglstIcon.ListImages(6).Picture
                Case "1"
                    Set .CellPicture = imglstIcon.ListImages(2).Picture
                Case "2"
                    Set .CellPicture = imglstIcon.ListImages(3).Picture
                Case "3"
                    Set .CellPicture = imglstIcon.ListImages(4).Picture
                Case "4"
                    Set .CellPicture = imglstIcon.ListImages(8).Picture
                Case "5"
                    Set .CellPicture = imglstIcon.ListImages(7).Picture
                Case "6"
                    Set .CellPicture = imglstIcon.ListImages(9).Picture
            End Select
            
            .CellPictureAlignment = 4
        Next
        
        For intRowCtr = 1 To .Rows - 1
            .RowHeight(intRowCtr) = 350
        Next
        
        .TextArray(1) = "Exam Code"
        .TextArray(2) = "Instrument ID"
        .TextArray(3) = "Section"
        .TextArray(4) = "Request ID"
        .TextArray(5) = "Section ID"
        .TextArray(6) = "Status"
        .TextArray(7) = "Sample ID"
        .TextArray(8) = "Specimen"
        .TextArray(9) = "Exam Name"
        .TextArray(10) = "RVID"
        
        .ColWidth(0) = 600
        .ColWidth(1) = 1000
        .ColWidth(3) = 2000
        .ColWidth(7) = 2000
        .ColWidth(8) = 1500
        .ColWidth(9) = 4170
        
        If rsTest.RecordCount > 0 Then
            For intRowCtr = 1 To .Rows - 1
                AddExamDet intRowCtr
            Next
        End If
        
    End With
        
End Sub

Private Sub AddExamDet(introw As Integer)
    
    Dim RowCtr As Integer
    
    With msfgResult
        Select Case .TextMatrix(introw, 13)
            Case "P"
                strSQLStatement = "EXEC sp_getParamProfileExam_Select '" & .TextMatrix(introw, 2) & "','" & _
                                    .TextMatrix(introw, 1) & "'"
                                    
                Set rsTest = myConnection.Execute(strSQLStatement)
                
                If rsTest.RecordCount > 0 Then
                    For RowCtr = 0 To rsTest.RecordCount - 1
                        msfgExamDet.TextMatrix(msfgExamDet.Rows - 1, 0) = rsTest.Fields(0)
                        msfgExamDet.TextMatrix(msfgExamDet.Rows - 1, 1) = rsTest.Fields(11)
                        msfgExamDet.TextMatrix(msfgExamDet.Rows - 1, 2) = rsTest.Fields(12)
                        msfgExamDet.TextMatrix(msfgExamDet.Rows - 1, 3) = rsTest.Fields(1)
                        rsTest.MoveNext
                        msfgExamDet.Rows = msfgExamDet.Rows + 1
                    Next
                End If
            Case "T"
                msfgExamDet.TextMatrix(msfgExamDet.Rows - 1, 0) = .TextMatrix(introw, 2)
                msfgExamDet.TextMatrix(msfgExamDet.Rows - 1, 1) = .TextMatrix(introw, 12)
                msfgExamDet.TextMatrix(msfgExamDet.Rows - 1, 2) = .TextMatrix(introw, 9)
                msfgExamDet.TextMatrix(msfgExamDet.Rows - 1, 3) = .TextMatrix(introw, 1)
                msfgExamDet.Rows = msfgExamDet.Rows + 1
        End Select
    End With
    
End Sub

Private Sub msfgResult_Click()
    With msfgResult
        If .Col = 1 And (Len(.TextMatrix(.Row, 7)) = 0) And _
            (Len(.TextMatrix(.Row, 0)) = 0 Or .TextMatrix(.Row, 0) = "0") And _
            Len(.TextMatrix(.Row, 10)) = 0 Then
            .CellAlignment = 0
            txtmsfgExam.Move .CellLeft + .Left, .CellTop + .Top, .CellWidth, .CellHeight
            txtmsfgExam.Visible = True
            txtmsfgExam.Enabled = True
            txtmsfgExam.Text = .TextMatrix(.Row, .Col)
            txtmsfgExam.SetFocus
        Else
            txtmsfgExam.Visible = False
            txtmsfgExam.Enabled = False
        End If
    End With
End Sub

Private Sub msfgResult_EnterCell()
    msfgResult_Click
End Sub

Private Sub msfgResult_LeaveCell()
    With msfgResult
        If (Len(.TextMatrix(.Row, 7)) = 0 Or .TextMatrix(.Row, 7) = "NEWEXAM") And _
            Len(.TextMatrix(.Row, 0)) = 0 Then
            .TextMatrix(.Row, .Col) = txtmsfgExam.Text
        End If
    End With
End Sub

Private Sub txtmsfgExam_Change()
    txtmsfgExam.Text = UCase(txtmsfgExam.Text)
    txtmsfgExam.SelStart = Len(txtmsfgExam.Text)
End Sub

Private Sub txtmsfgExam_GotFocus()
    highLight
End Sub

Private Sub txtmsfgExam_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSelectionValue As String
    If KeyCode = 113 Then
        strSelectionValue = showfrmExamSelection(txtmsfgExam.Text, "ALL")
        If Len(strSelectionValue) > 0 Then
'0=status
'1=examcode
'2=instrumentid
'3=rvid
'4=requestid
'5=sectionid
'6=status
'7=sampleid
'8=specimen
'9=examname
            If Not blnExamOccured(strSelectionValue) Then
                With msfgResult
                    .TextMatrix(.Row, 0) = "0"
                    .CellAlignment = 1
                    .TextMatrix(.Row, 1) = getValue(strSelectionValue, 1)
                    .TextMatrix(.Row, 2) = getValue(strSelectionValue, 4)
                    .TextMatrix(.Row, 3) = getValue(strSelectionValue, 11)
                    .TextMatrix(.Row, 4) = RequestID
                    .TextMatrix(.Row, 5) = getValue(strSelectionValue, 10)
                    .TextMatrix(.Row, 6) = "0"
                    .TextMatrix(.Row, 7) = ""
                    .TextMatrix(.Row, 8) = getValue(strSelectionValue, 3)
                    .TextMatrix(.Row, 9) = getValue(strSelectionValue, 2)
                    .TextMatrix(.Row, 10) = ""
                    .TextMatrix(.Row, 11) = getValue(strSelectionValue, 5)
                    .TextMatrix(.Row, 12) = getValue(strSelectionValue, 6)
                    .TextMatrix(.Row, 13) = getValue(strSelectionValue, 7)
                    .TextMatrix(.Row, 14) = getValue(strSelectionValue, 8)
                    .TextMatrix(.Row, 15) = getValue(strSelectionValue, 9)
                    .Col = 0
                    .CellForeColor = vbWhite
                    
                    AddExamDet .Row
                    
                    If Len(.TextMatrix(.Rows - 1, 0)) > 0 Then
                        .Rows = .Rows + 1
                    End If
                    .Row = .Rows - 1

                    .Col = 0
                    .CellForeColor = vbWhite
                    .CellBackColor = vbWhite
                    
                    Set .CellPicture = imglstIcon.ListImages(6).Picture
                    .CellPictureAlignment = 4
                    .Col = 1
                    
                    .RowHeight(.Row) = 350
                    txtmsfgExam.Move .Left + .CellLeft, .Top + .CellTop
                    txtmsfgExam.Visible = True
                    txtmsfgExam.Text = ""
                    .CellAlignment = 0
                End With
            Else
                MsgBox "Unable to Add " & getValue(strSelectionValue, 2)
            End If
        End If
        If txtmsfgExam.Enabled Then
            txtmsfgExam.SetFocus
        End If
    ElseIf KeyCode = 38 Then
        With msfgResult
            .TopRow = .Row
            If .Row > 1 Then
                msfgResult_LeaveCell
                .Row = .Row - 1
                .SetFocus
                .ColSel = .Cols - 1
                msfgResult_EnterCell
                highLight
            End If
        End With
    ElseIf KeyCode = 40 Then
        With msfgResult
             .TopRow = .Row
            If .Row < .Rows - 1 Then
                msfgResult_LeaveCell
                .Row = .Row + 1
                .SetFocus
                msfgResult_EnterCell
                highLight
            End If
        End With
    ElseIf KeyCode = vbKeyEscape Then
        txtmsfgExam.Visible = False
    Else
        If Len(msfgResult.TextMatrix(msfgResult.Row, 1)) > 0 Then
            DelExamDet msfgResult.TextMatrix(msfgResult.Row, 1)
        End If
        With msfgResult
            If Len(msfgResult.TextMatrix(msfgResult.Row, 1)) > 0 Then
                If .Rows > 2 Then
                    .RemoveItem (msfgResult.Row)
                Else
                    For intRowCtr = 0 To msfgResult.Cols - 1
                        msfgResult.TextMatrix(msfgResult.Row, intRowCtr) = ""
                    Next
                End If
                MsgBox .TextMatrix(.Row, 1)
            End If
        End With
    End If
End Sub

Private Function blnExamOccured(examCriteria As String) As Boolean
    Dim tempInsID As String
    Dim tempExamCode As String
    Dim examType As String
    Dim rsCheckParam As New ADODB.Recordset
    Dim msfgExamDetRowCtr As Integer
    Dim rsCheckParamRowCtr As Integer
    
    tempInsID = getValue(examCriteria, 4)
    tempExamCode = getValue(examCriteria, 1)
    examType = getValue(examCriteria, 7)
    
    Select Case examType
        Case "P"
            strSQLStatement = "EXEC sp_getParamProfileExam_Select '" & tempInsID & "','" & tempExamCode & "'"
            
            Set rsCheckParam = myConnection.Execute(strSQLStatement)
            
            If rsCheckParam.RecordCount > 0 Then
                With msfgExamDet
                    If Len(.TextMatrix(1, 0)) > 0 Then
                        For rsCheckParamRowCtr = 0 To rsCheckParam.RecordCount - 1
                            For msfgExamDetRowCtr = 1 To .Rows - 1
                                If RTrim(rsCheckParam.Fields(0).Value) = RTrim(.TextMatrix(msfgExamDetRowCtr, 0)) And _
                                    RTrim(rsCheckParam.Fields(11).Value) = RTrim(.TextMatrix(msfgExamDetRowCtr, 1)) Then
                                    blnExamOccured = True
                                    Exit For
                                End If
                            Next
                            If blnExamOccured Then
                                Exit For
                            Else
                                rsCheckParam.MoveNext
                            End If
                        Next
                    Else
                        blnExamOccured = False
                    End If
                    rsCheckParam.Close
                End With
            Else
                blnExamOccured = False
            End If
        Case "T"
            With msfgExamDet
                If Len(.TextMatrix(1, 0)) > 0 Then
                    tempExamCode = getValue(examCriteria, 6)
                    For msfgExamDetRowCtr = 1 To .Rows - 1
                        If tempInsID = RTrim(.TextMatrix(msfgExamDetRowCtr, 0)) And _
                            tempExamCode = RTrim(.TextMatrix(msfgExamDetRowCtr, 1)) Then
                            blnExamOccured = True
                            Exit For
                        End If
                    Next
                Else
                    blnExamOccured = False
                End If
            End With
    End Select
    
End Function

Private Sub DelExamDet(reqExamCode As Integer)
    
    Dim delRowCtr As Integer

    With msfgExamDet
        delRowCtr = 1
        Do
            If .TextMatrix(delRowCtr, 3) = reqExamCode Then
                .RemoveItem (delRowCtr)
            Else
                delRowCtr = delRowCtr + 1
            End If
        Loop Until delRowCtr = .Rows - 1
    End With

End Sub

Private Sub txtmsfgExam_KeyPress(KeyAscii As Integer)
'    Dim strSelectionValue As String
'
'    If KeyAscii = 13 Then
'        If Len(Trim(txtmsfgExam.Text)) = 0 Then
'            If getValue(1, getUserDet(strRequestBy)) = 4 Then
'                strSelectionValue = showfrmExamSelection(txtmsfgExam.Text, "ALL")
'            Else
'                strSelectionValue = showfrmExamSelection(txtmsfgExam.Text)
'            End If
'        Else
'            If getValue(1, getUserDet(strRequestBy)) = 4 Then
'                strSQLStatement = "EXEC sp_ReqExam_Select NULL," & _
'                                                "NULL,'" & Replace(txtmsfgExam.Text, "'", "''") & "'"
'            Else
'                strSQLStatement = "EXEC sp_ReqExam_Select " & IIf(Len(InstrumentID) > 0, "'" & InstrumentID & "'", "NULL") & "," & _
'                                                "NULL,'" & Replace(txtmsfgExam.Text, "'", "''") & "'"
'            End If
'
'            Set rsTest = myConnection.Execute(strSQLStatement)
'            If rsTest.RecordCount > 0 Then
'                If rsTest.RecordCount > 1 Then
'                    If getValue(1, getUserDet(strRequestBy)) = 4 Then
'                        strSelectionValue = showfrmExamSelection(txtmsfgExam.Text, "ALL")
'                    Else
'                        strSelectionValue = showfrmExamSelection(txtmsfgExam.Text)
'                    End If
'                Else
'                    For intColCtr = 0 To rsTest.Fields.Count - 1
'                        If Len(strSelectionValue) = 0 Then
'                            strSelectionValue = "|" & rsTest.Fields(intColCtr) & "|"
'                        Else
'                            strSelectionValue = strSelectionValue & rsTest.Fields(intColCtr) & "|"
'                        End If
'                    Next
'                End If
'            Else
'                strSelectionValue = showfrmExamSelection(txtmsfgExam.Text)
'            End If
'        End If
'        If Len(strSelectionValue) > 0 Then
'            If Not blnExamOccured(strSelectionValue) Then
'                With msfgResult
'                    .TextMatrix(.Row, 0) = "0"
'                    .CellAlignment = 1
'                    .TextMatrix(.Row, 1) = getValue(1, strSelectionValue)
'                    .TextMatrix(.Row, 2) = getValue(4, strSelectionValue)
'                    .TextMatrix(.Row, 3) = getValue(11, strSelectionValue)
'                    .TextMatrix(.Row, 4) = RequestID
'                    .TextMatrix(.Row, 5) = getValue(10, strSelectionValue)
'                    .TextMatrix(.Row, 6) = "0"
'                    .TextMatrix(.Row, 7) = ""
'                    .TextMatrix(.Row, 8) = getValue(3, strSelectionValue)
'                    .TextMatrix(.Row, 9) = getValue(2, strSelectionValue)
'                    .TextMatrix(.Row, 10) = ""
'                    .TextMatrix(.Row, 11) = getValue(5, strSelectionValue)
'                    .TextMatrix(.Row, 12) = getValue(6, strSelectionValue)
'                    .TextMatrix(.Row, 13) = getValue(7, strSelectionValue)
'                    .TextMatrix(.Row, 14) = getValue(8, strSelectionValue)
'                    .TextMatrix(.Row, 15) = getValue(9, strSelectionValue)
'                    .Col = 0
'                    .CellForeColor = vbWhite
'
'                    AddExamDet .Row
'
'                    If Len(.TextMatrix(.Rows - 1, 0)) > 0 Then
'                        .Rows = .Rows + 1
'                    End If
'
'                    .Row = .Rows - 1
'
'                    .Col = 0
'                    .CellBackColor = vbWhite
'                    .CellForeColor = vbBlack
'
'                    Set .CellPicture = imglstIcon.ListImages(6).Picture
'                    .CellPictureAlignment = 4
'                    .Col = 1
'
'                    .RowHeight(.Row) = 350
'                    txtmsfgExam.Move .Left + .CellLeft, .Top + .CellTop
'                    txtmsfgExam.Visible = True
'                    txtmsfgExam.Text = ""
'                    .CellAlignment = 0
'
'                End With
'            Else
'                MsgBox "Unable to Add " & getValue(2, strSelectionValue)
'            End If
'            KeyAscii = 0
'        End If
'        If txtmsfgExam.Enabled Or txtmsfgExam.Visible Then
'            txtmsfgExam.SetFocus
'        End If
'    End If
End Sub
