VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmProcess 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Result Processing"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6435
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1950
      TabIndex        =   0
      Top             =   3060
      Width           =   6375
      Begin VB.Timer Timer1 
         Interval        =   600
         Left            =   5790
         Top             =   120
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result Processing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   180
         Width           =   1515
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   6375
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   345
         Left            =   5160
         TabIndex        =   6
         Top             =   600
         Width           =   1125
      End
      Begin MSComctlLib.ProgressBar progress 
         Height          =   285
         Left            =   90
         TabIndex        =   3
         Top             =   210
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSComctlLib.ProgressBar pbSub 
         Height          =   75
         Left            =   90
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   132
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Searching..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SaveResultFailed As Boolean  'set by SaveResult when the result could not be committed
Private rstExam2 As New ADODB.Recordset
Private rstUnprocessed As New ADODB.Recordset
Private ResultOtherInfo() As String
Private Cancelled As Boolean

Private Function SaveResult(msgHL7format As Variant) As String
Dim ctr1 As Integer
Dim counter As Integer
Dim myResults As Variant
Dim myResults2 As Variant
Dim strSql As String
Dim rstTemp As ADODB.Recordset
Dim SampleNum As String
Dim rstSearch As New ADODB.Recordset
Dim temps As String
Dim tempResultNum As String
Dim myResultValue As String
Dim SaveHeader As Boolean
Dim myTempTime As String
Dim tPatientName, tSpecimenID, tDate As String
Dim hasBeginTrans As Boolean
Dim errNum As Long
Dim errDesc As String
SaveHeader = False
SaveResultFailed = False
tPatientName = ""
tSpecimenID = ""
tDate = ""
On Error GoTo errFound
myTempTime = Format(Now, "HH:mm:ss")
myResults = Split(msgHL7format, vbCr)
counter = UBound(myResults) - 1
If rstExam2.RecordCount > 0 Then
rstExam2.MoveFirst
hasBeginTrans = False
myConnection.BeginTrans
hasBeginTrans = True
For ctr1 = 0 To counter
    myResults2 = Split(myResults(ctr1), "|")
    If Len(ResultOtherInfo(3)) > 3 Then SampleNum = ResultOtherInfo(3) Else SampleNum = String(3 - Len(ResultOtherInfo(3)), "0") & ResultOtherInfo(3)
        If ProcessDateOnly = 1 Then
            If DBServerType = "MYSQL SERVER" Then
                temps = "Select * from tbLabInstrumentResults where InstrumentID = '" & _
                    InstrumentID & "' and PatientName = '" & ResultOtherInfo(2) & "' and AccessionNumOrg = '" & _
                    ResultOtherInfo(6) & "' and format(Resultdate,'mm/dd/yyyy') = '" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy") & "' and ExamName = '" & myResults2(3) & "'"
                SaveResult = formatDT(myResults2(14), 3)
            Else
                temps = "Select * from tbLabInstrumentResults where InstrumentID = '" & _
                    InstrumentID & "' and PatientName = '" & ResultOtherInfo(2) & "' and AccessionNumOrg = '" & _
                    ResultOtherInfo(6) & "' and format(Resultdate,'mm/dd/yyyy') = '" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy") & "' and ExamName = '" & myResults2(3) & "'"
                SaveResult = formatDT(myResults2(14), 3)
            End If
        Else
            If DBServerType = "MYSQL SERVER" Then
                temps = "Select * from tbLabInstrumentResults where InstrumentID = '" & _
                    InstrumentID & "' and PatientName = '" & ResultOtherInfo(2) & "' and AccessionNumOrg = '" & _
                    ResultOtherInfo(6) & "' and Resultdate = '" & Format(formatDT(myResults2(14), 3), "yyyy-mm-dd HH:mm:ss") & "' and ExamName = '" & myResults2(3) & "'"
                SaveResult = formatDT(myResults2(14), 3)
            Else
                
                If Val(GetSetting("Lab-I", "Settings", "UseCurrentTime", "0")) = 1 Then
                    temps = "Select a.*, b.* from tbLabInstrumentResults a, tbLabInstrumentRValues b where a.ResultID = b.ResultID and a.InstrumentID = '" & _
                        InstrumentID & "' and a.PatientName = '" & ResultOtherInfo(2) & "' and a.AccessionNumOrg = '" & _
                        ResultOtherInfo(6) & "' and format(a.Resultdate,'mm/dd/yyyy HH:mm') = '" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy ") & myTempTime & "' and b.ExamName = '" & myResults2(3) & "'"
                    
                        SaveResult = formatDT(myResults2(14), 3)
                Else
                    temps = "Select a.*, b.* from tbLabInstrumentResults a, tbLabInstrumentRValues b where a.ResultID = b.ResultID and a.InstrumentID = '" & _
                        InstrumentID & "' and a.PatientName = '" & ResultOtherInfo(2) & "' and a.AccessionNumOrg = '" & _
                        ResultOtherInfo(6) & "' and format(a.Resultdate,'mm/dd/yyyy HH:mm') = '" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy HH:mm") & "' and b.ExamName = '" & myResults2(3) & "'"
                    
                    SaveResult = formatDT(myResults2(14), 3)
                End If
            End If
        End If
    Set rstSearch = myConnection.Execute(temps)
    'If saveStringResult Then myResultValue = "'" & IIf(Len(Trim(myResults2(5))) = 0, "0", myResults2(5)) & "'" Else myResultValue = IIf(Len(Trim(myResults2(5))) = 0, "0", myResults2(5))
    myResultValue = "'" & IIf(Len(Trim(myResults2(5))) = 0, "0", myResults2(5)) & "'"
    If rstSearch.RecordCount > 0 Then
        If ProcessDateOnly = 1 Then
            If DBServerType = "MYSQL SERVER" Then
                temps = "Update tbLabInstrumentResults SET Result = " & myResultValue & " " & vbCrLf & _
                    "where Status = '0' and InstrumentID = '" & _
                    InstrumentID & "' and PatientName = '" & ResultOtherInfo(2) & "' and AccessionNumOrg = '" & _
                    ResultOtherInfo(6) & "' and format(Resultdate,'mm/dd/yyyy') = #" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy") & "# and ExamName = '" & myResults2(3) & "'"
                SaveResult = formatDT(myResults2(14), 3)
            Else
                temps = "Update tbLabInstrumentResults SET Result = " & myResultValue & " " & vbCrLf & _
                    "where Status = '0' and InstrumentID = '" & _
                    InstrumentID & "' and PatientName = '" & ResultOtherInfo(2) & "' and AccessionNumOrg = '" & _
                    ResultOtherInfo(6) & "' and format(Resultdate,'mm/dd/yyyy') = #" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy") & "# and ExamName = '" & myResults2(3) & "'"
                SaveResult = formatDT(myResults2(14), 3)
            End If
        Else
            If DBServerType = "MYSQL SERVER" Then
                temps = "Update tbLabInstrumentResults SET Result = " & myResultValue & " " & vbCrLf & _
                    "where Status = '0' and InstrumentID = '" & _
                    InstrumentID & "' and PatientName = '" & ResultOtherInfo(2) & "' and AccessionNumOrg = '" & _
                    ResultOtherInfo(6) & "' and Resultdate = '" & Format(formatDT(myResults2(14), 3), "yyyy-mm-dd HH:mm:ss") & "' and ExamName = '" & myResults2(3) & "'"
                SaveResult = formatDT(myResults2(14), 3)
            Else
                temps = "Update tbLabInstrumentResults SET Result = " & myResultValue & " " & vbCrLf & _
                    "where Status = '0' and InstrumentID = '" & _
                    InstrumentID & "' and PatientName = '" & ResultOtherInfo(2) & "' and AccessionNumOrg = '" & _
                    ResultOtherInfo(6) & "' and Resultdate = #" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy HH:mm:ss") & "# and ExamName = '" & myResults2(3) & "'"
                SaveResult = formatDT(myResults2(14), 3)
            End If
        End If
    Else
        If ProcessDateOnly = 1 Then
            If DBServerType = "MYSQL SERVER" Then
                temps = "Insert into tbLabInstrumentResults (InstrumentID, Resultdate, Accessionnum, PatientName, ExamName, Result, Unit, Status, ProfileExam, AccessionnumOrg, IdNum, BirthDate, SortNum, Room, Age, Sex, Specimen, Doctor)" & vbCrLf & _
                "Select '" & InstrumentID & "',#" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy") & "#,'" & ResultOtherInfo(6) & "','" & ResultOtherInfo(2) & "','" & myResults2(3) & "'," & myResultValue & ",'" & myResults2(6) & "','0','" & DefaultProfileID & "','" & ResultOtherInfo(6) & "','" & ResultOtherInfo(7) & "','" & _
                ResultOtherInfo(8) & "'," & myResults2(1) & ",'" & ResultOtherInfo(9) & "','" & ResultOtherInfo(10) & "','" & ResultOtherInfo(11) & "','" & ResultOtherInfo(12) & "','" & ResultOtherInfo(14) & "';"
                
                SaveResult = formatDT(myResults2(14), 3)
            Else
                temps = "Insert into tbLabInstrumentResults (InstrumentID, Resultdate, Accessionnum, PatientName, ExamName, Result, Unit, Status, ProfileExam, AccessionnumOrg, IdNum, BirthDate, SortNum, Room, Age, Sex, Specimen, Doctor)" & vbCrLf & _
                "Select '" & InstrumentID & "',#" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy") & "#,'" & ResultOtherInfo(6) & "','" & ResultOtherInfo(2) & "','" & myResults2(3) & "'," & myResultValue & ",'" & myResults2(6) & "','0','" & DefaultProfileID & "','" & ResultOtherInfo(6) & "','" & ResultOtherInfo(7) & "','" & _
                ResultOtherInfo(8) & "'," & myResults2(1) & ",'" & ResultOtherInfo(9) & "','" & ResultOtherInfo(10) & "','" & ResultOtherInfo(11) & "','" & ResultOtherInfo(12) & "','" & ResultOtherInfo(14) & "';"
                
                SaveResult = formatDT(myResults2(14), 3)
            End If
        Else
            If DBServerType = "MYSQL SERVER" Then
                temps = "Insert into tbLabInstrumentResults (InstrumentID, Resultdate, Accessionnum, PatientName, ExamName, Result, Unit, Status, ProfileExam, AccessionnumOrg, IdNum, BirthDate, SortNum, Room, Age, Sex, Specimen, Doctor)" & vbCrLf & _
                "Select '" & InstrumentID & "','" & Format(formatDT(myResults2(14), 3), "yyyy-mm-dd HH:mm:ss") & "','" & ResultOtherInfo(6) & "','" & ResultOtherInfo(2) & "','" & myResults2(3) & "'," & myResultValue & ",'" & myResults2(6) & "','0','" & DefaultProfileID & "','" & ResultOtherInfo(6) & "','" & ResultOtherInfo(7) & "','" & _
                ResultOtherInfo(8) & "'," & myResults2(1) & ",'" & ResultOtherInfo(9) & "','" & ResultOtherInfo(10) & "','" & ResultOtherInfo(11) & "','" & ResultOtherInfo(12) & "','" & ResultOtherInfo(14) & "';"
                
                SaveResult = formatDT(myResults2(14), 3)
            Else
                If Not SaveHeader Then
                    If Len(Trim(ResultOtherInfo(2))) = 0 Then
                        tPatientName = ResultOtherInfo(6)
                    Else
                        tPatientName = ResultOtherInfo(2)
                    End If
                    If Val(GetSetting("Lab-I", "Settings", "UseCurrentTime", "0")) = 1 Then
                        temps = "Insert into tbLabInstrumentResults (InstrumentID, Resultdate, Accessionnum, PatientName, Status, ProfileExam, AccessionnumOrg, IdNum, BirthDate, Room, Age, Sex, Specimen, Doctor, ResultID,Operator,Comments)" & vbCrLf & _
                        "Select '" & InstrumentID & "',#" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy ") & myTempTime & "#,'" & ResultOtherInfo(6) & "','" & tPatientName & "','0','" & DefaultProfileID & "','" & ResultOtherInfo(6) & "','" & ResultOtherInfo(7) & "','" & _
                        ResultOtherInfo(8) & "','" & ResultOtherInfo(9) & "','" & ResultOtherInfo(10) & "','" & ResultOtherInfo(11) & "','" & ResultOtherInfo(12) & "','" & ResultOtherInfo(14) & "','" & GenerateResultID(InstrumentID & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy ") & myTempTime & ResultOtherInfo(6) & ResultOtherInfo(2)) & "','" & ResultOtherInfo(17) & "','" & ResultOtherInfo(18) & "';"
                        tPatientName = ResultOtherInfo(2)
                        tSpecimenID = ResultOtherInfo(6)
                        tDate = Format(formatDT(myResults2(14), 3), "mm/dd/yyyy ") & myTempTime
                    Else
                        temps = "Insert into tbLabInstrumentResults (InstrumentID, Resultdate, Accessionnum, PatientName, Status, ProfileExam, AccessionnumOrg, IdNum, BirthDate, Room, Age, Sex, Specimen, Doctor, ResultID,Operator, Comments)" & vbCrLf & _
                        "Select '" & InstrumentID & "',#" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy HH:mm:ss") & "#,'" & ResultOtherInfo(6) & "','" & tPatientName & "','0','" & DefaultProfileID & "','" & ResultOtherInfo(6) & "','" & ResultOtherInfo(7) & "','" & _
                        ResultOtherInfo(8) & "','" & ResultOtherInfo(9) & "','" & ResultOtherInfo(10) & "','" & ResultOtherInfo(11) & "','" & ResultOtherInfo(12) & "','" & ResultOtherInfo(14) & "','" & GenerateResultID(InstrumentID & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy HH:mm:ss") & ResultOtherInfo(6) & ResultOtherInfo(2)) & "','" & ResultOtherInfo(17) & "','" & ResultOtherInfo(18) & "';"
                        tPatientName = ResultOtherInfo(2)
                        tSpecimenID = ResultOtherInfo(6)
                        tDate = Format(formatDT(myResults2(14), 3), "mm/dd/yyyy HH:mm:ss")
                    End If
                    myConnection.Execute temps
                    SaveHeader = True
                End If
                
                If Val(GetSetting("Lab-I", "Settings", "UseCurrentTime", "0")) = 1 Then
                    temps = "Insert into tbLabInstrumentRValues (ExamName, Result, Unit, SortNum, ResultID)" & vbCrLf & _
                    "Select '" & myResults2(3) & "'," & myResultValue & ",'" & myResults2(6) & "'," & myResults2(1) & ",'" & GenerateResultID(InstrumentID & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy ") & myTempTime & ResultOtherInfo(6) & ResultOtherInfo(2)) & "';"
                    
                    SaveResult = formatDT(myResults2(14), 3)
                Else
                    temps = "Insert into tbLabInstrumentRValues (ExamName, Result, Unit, SortNum, ResultID)" & vbCrLf & _
                    "Select '" & myResults2(3) & "'," & myResultValue & ",'" & myResults2(6) & "'," & myResults2(1) & ",'" & GenerateResultID(InstrumentID & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy HH:mm:ss") & ResultOtherInfo(6) & ResultOtherInfo(2)) & "';"
                    
                    SaveResult = formatDT(myResults2(14), 3)
                End If
            End If
        End If
    End If
    If Val(GetSetting("HostForms_SA", "Settings", "SaveStringResult", "0")) = 1 Then
        myConnection.Execute temps
    Else
        If Len(Trim(myResults2(5))) > 0 And IsNumeric(myResults2(5)) And (Len(Trim(ResultOtherInfo(2))) > 0 Or Len(Trim(ResultOtherInfo(3))) > 0) Then myConnection.Execute temps
    End If
Set rstSearch = Nothing
Next ctr1
If Autoprint Then If Len(tPatientName) Then myConnection.Execute "Insert Into LabTempPrint (InstrumentID, SpecimenID, PatientName, Resultdate) select '" & InstrumentID & "','" & tSpecimenID & "','" & tPatientName & "','" & tDate & "'"
myConnection.CommitTrans
hasBeginTrans = False
End If
errFound:
    If Err.Number <> 0 Then
        errNum = Err.Number
        errDesc = Err.Description
        SaveResultFailed = True
        If hasBeginTrans Then
            'Roll back on every error (including 9) so no transaction is left open
            hasBeginTrans = False
            On Error Resume Next
            myConnection.RollbackTrans
            On Error GoTo 0
            SaveSetting "Lab-I", "Database", "ProcessError", "Yes"
        End If
        Select Case errNum
            Case 9
                MsgBox "Error in conversion. Unable to process data.", vbOKOnly + vbCritical, "Conversion Error"
            Case Else
                MsgBox errNum & " " & errDesc, vbCritical + vbOKOnly, "Error"
        End Select
    End If
End Function

Private Sub processData()
Dim ctr As Integer
Dim ctrOBX As Integer
Dim ctrOBI As Integer
Dim arrOBX As Variant
Dim arrOBI As Variant
Dim tempstr As String
Dim mHL7format As String
Dim myDate As String
Dim rowFailed As Boolean
    If DBServerType = "LOCAL DATABASE" Then
        Set rstExam2 = myConnection.Execute("select * from tbcolabexam where Active = True and InstrumentID = '" & InstrumentID & "' and HisExamCode <> '0' and HisResultCode <> '0'")
    ElseIf DBServerType = "MSSQL SERVER" Then
        Set rstExam2 = myConnection.Execute("select * from tbcolabexam where Active = 1 and InstrumentID = '" & InstrumentID & "' and HisExamCode <> '0' and HisResultCode <> '0'")
    End If
    Set rstUnprocessed = myConnection.Execute("select * from labunprocessed where instrumentID = '" & InstrumentID & "'")
    Select Case MessageFormat
    Case "STKS"
        If rstUnprocessed.RecordCount > 0 Then
            progress.Value = 1
            progress.Min = 1
            progress.Max = rstUnprocessed.RecordCount + 1
            rstUnprocessed.MoveFirst
            Do While Not rstUnprocessed.EOF
                DoEvents
                lblStatus.Caption = "Processing " & progress.Value & " of " & progress.Max - 1 & " record(s)..."
                tempstr = Replace(rstUnprocessed.Fields("resultdata").Value, "'", "''")
                ResultOtherInfo = get_OtherResultInfo(tempstr)
                mHL7format = segment_OBX(tempstr, rstExam2)
                myDate = SaveResult(mHL7format)
                rowFailed = SaveResultFailed
                progress.Value = progress.Value + 1
                
                If EnableDTP = 1 And Val(GetSetting("Lab-I", "Settings", "Autoprocess", "0")) = 0 Then
                    ObjDTP.ProcessDTP segment_OBX2(tempstr, rstExam2), get_OtherResultInfo(tempstr, 1), tempstr
                    If ObjDTP.ErrorNum <> 0 Then
'                        rtbRcvDisplay.Text = rtbRcvDisplay.Text & ObjDTP.ErrorDesc
                    End If
                End If

                
                'Keep the raw message if any part of it failed to save, so it can be reprocessed
                If Not rowFailed Then myConnection.Execute "Delete from labunprocessed where ResultID = " & rstUnprocessed.Fields("ResultID").Value & " and InstrumentID = '" & InstrumentID & "'"
                rstUnprocessed.MoveNext
                If Cancelled Then Exit Do
            Loop
        End If
    Case Else
        If rstUnprocessed.RecordCount > 0 Then
            progress.Value = 1
            progress.Min = 1
            progress.Max = rstUnprocessed.RecordCount + 1
            rstUnprocessed.MoveFirst
            Do While Not rstUnprocessed.EOF
                DoEvents
                lblStatus.Caption = "Processing " & progress.Value & " of " & progress.Max - 1 & " record(s)..."
                tempstr = Replace(rstUnprocessed.Fields("resultdata").Value, "'", "''")
                rowFailed = False
                arrOBX = Split(MsgDef.segment_OBX(tempstr), "<SP>")
                arrOBI = Split(MsgDef.segment_oBI(tempstr), vbCrLf)
                ctrOBX = UBound(arrOBX)
                ctrOBI = UBound(arrOBI)
                If ctrOBX = ctrOBI Then
                    If ctrOBX >= 1 Then
                        pbSub.Visible = True
                        pbSub.Max = ctrOBX
                        pbSub.Min = 0
                        For ctr = 0 To ctrOBX
                            ResultOtherInfo = Split(arrOBI(ctr), "|")
                            If MsgDef.segment_oBI(tempstr) <> "ERRCONV" Then
                                SaveResult arrOBX(ctr)
                                If SaveResultFailed Then rowFailed = True
                            End If
                            If ctr > 0 Then pbSub.Value = ctr
                            
                            If EnableDTP = 1 And Val(GetSetting("Lab-I", "Settings", "Autoprocess", "0")) = 0 Then
                                ObjDTP.ProcessDTP arrOBX(ctr), ResultOtherInfo
                                If ObjDTP.ErrorNum <> 0 Then
                                    'rtbRcvDisplay.Text = rtbRcvDisplay.Text & ObjDTP.ErrorDesc
                                End If
                            End If

                        Next ctr
                        pbSub.Visible = False
                    Else
                        pbSub.Visible = False
                        ResultOtherInfo = Split(MsgDef.segment_oBI(tempstr), "|")
                        If MsgDef.segment_oBI(tempstr) <> "ERRCONV" Then
                            SaveResult MsgDef.segment_OBX(tempstr)
                            If SaveResultFailed Then rowFailed = True
                            
                            If EnableDTP = 1 And Val(GetSetting("Lab-I", "Settings", "Autoprocess", "0")) = 0 Then
                                ObjDTP.ProcessDTP MsgDef.segment_OBX(tempstr), MsgDef.segment_oBI(tempstr), tempstr
                                If ObjDTP.ErrorNum <> 0 Then
                                    'rtbRcvDisplay.Text = rtbRcvDisplay.Text & ObjDTP.ErrorDesc
                                End If
                            End If

                        End If
                    End If
                End If
                progress.Value = progress.Value + 1
                'Keep the raw message if any part of it failed to save, so it can be reprocessed
                If Not rowFailed Then myConnection.Execute "Delete from labunprocessed where ResultID = " & rstUnprocessed.Fields("ResultID").Value & " and InstrumentID = '" & InstrumentID & "'"
                rstUnprocessed.MoveNext
                If Cancelled Then Exit Do
            Loop
        End If
    End Select
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
End Sub

Private Sub Form_Load()
setFormColors Me
Cancelled = False
lblStatus.Caption = "Searching..."
End Sub

Private Sub Timer1_Timer()
If frmProcess.Visible Then
    Timer1.Enabled = False
    Me.MousePointer = 11
    processData
    Me.MousePointer = 0
End If
lblStatus.Caption = "Searching..."
End Sub

