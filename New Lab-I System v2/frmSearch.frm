VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search"
   ClientHeight    =   3345
   ClientLeft      =   4320
   ClientTop       =   810
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
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
      Left            =   1590
      Style           =   2  'Dropdown List
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2805
   End
   Begin VB.CommandButton cmdClear 
      Cancel          =   -1  'True
      Caption         =   "Reset..."
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2490
      Width           =   1065
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   3270
      TabIndex        =   6
      Top             =   2490
      Width           =   1095
   End
   Begin VB.TextBox txtPatientName 
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
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1320
      Width           =   3315
   End
   Begin VB.TextBox txtSequenceNum 
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
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1680
      Width           =   3315
   End
   Begin VB.CheckBox chkResultDate 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   90
      TabIndex        =   0
      Top             =   300
      Width           =   195
   End
   Begin VB.ListBox lstPatientName 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4530
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New..."
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   2490
      Width           =   1035
   End
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5280
      TabIndex        =   9
      Top             =   510
      Width           =   1155
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
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
         Top             =   150
         Width           =   675
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   0
      Left            =   5370
      TabIndex        =   10
      Top             =   1920
      Width           =   2565
   End
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   345
      Left            =   1290
      TabIndex        =   2
      Top             =   630
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MM/dd/yyyy HH:mm"
      Format          =   114294787
      CurrentDate     =   38040
   End
   Begin MSComCtl2.DTPicker dtstart 
      Height          =   315
      Left            =   1290
      TabIndex        =   1
      Top             =   270
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MM/dd/yyyy HH:mm"
      Format          =   114294787
      CurrentDate     =   38040
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Result Date:"
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
      Index           =   4
      Left            =   360
      TabIndex        =   17
      Top             =   300
      Width           =   870
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
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
      Left            =   990
      TabIndex        =   16
      Top             =   690
      Width           =   225
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   4320
      X2              =   120
      Y1              =   1155
      Y2              =   1155
   End
   Begin VB.Line Line1 
      X1              =   4320
      X2              =   120
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Name:"
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
      Left            =   90
      TabIndex        =   15
      Top             =   1350
      Width           =   975
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sequence #:"
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
      TabIndex        =   14
      Top             =   1740
      Width           =   915
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Predefined Search:"
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
      Left            =   90
      TabIndex        =   13
      Top             =   2070
      Width           =   1395
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboStatus_Click()
If cboStatus.Visible And SearchReset = False Then
    Select Case cboStatus.Text
        Case "All", "Unedited", "Printed", "Saved"
            chkResultDate.Value = 0
            txtPatientName.Text = ""
            txtSequenceNum.Text = ""
        Case "All-Today"
            chkResultDate.Value = 1
            dtstart.Value = Format(Now, "mm/dd/yyyy") & " 00:00"
            dtEnd.Value = Format(Now, "mm/dd/yyyy") & " 23:59"
        Case "All-Yesterday"
        Case "Unedited-Today"
            chkResultDate.Value = 1
            dtstart.Value = Format(Now, "mm/dd/yyyy") & " 00:00"
            dtEnd.Value = Format(Now, "mm/dd/yyyy") & " 23:59"
        Case "Unedited-Yesterday"
        Case Else
    End Select
    cmdSearch_Click
End If
End Sub

Private Sub cboStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub chkResultDate_Click()
If chkResultDate.Value = 1 Then
    dtstart.Enabled = True
    dtEnd.Enabled = True
Else
    dtstart.Enabled = False
    dtEnd.Enabled = False
End If
End Sub

Private Sub chkResultDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub cmdClear_Click()
    Set rstResult = Nothing
    txtPatientName.Text = ""
    txtSequenceNum.Text = ""
    If ProcessDateOnly = 1 Then
        dtstart.Value = CDate(Format(Now, "mm/dd/yyyy 00:00"))
        dtEnd.Value = Now
    Else
        dtstart.Value = CDate(Format(Now, "mm/dd/yyyy 00:00"))
        dtEnd.Value = CDate(Format(Now, "mm/dd/yyyy 23:59"))
    End If
    dtstart.Enabled = True
    dtEnd.Enabled = True
    chkResultDate.Value = 1
    If chkResultDate.Visible Then chkResultDate.SetFocus
    Unload frmSearchResult
    isResultVisible = False
    SearchReset = True
    cboStatus.ListIndex = 0
    SearchReset = False
End Sub

Private Sub cmdNew_Click()
    HostSA.showResultEntry InstrumentID, "NEW", "NEW", "NEW", 1, 1, 1
End Sub

Public Sub cmdSearch_Click()
Dim rstTemp As ADODB.Recordset
Dim CodeLoc As String
On Error GoTo errFound
Me.MousePointer = 11
If Val(GetSetting("Lab-I", "Others", "ENABLEWARNING", "0")) = 1 Then
    CodeLoc = "GETWARNING"
    Set rstTemp = myConnection.Execute("Select Distinct Resultdate as [Result Date], PatientName as [Patient Name], Accessionnum as [Sequence #],val(Accessionnum), refnum, patientname, newpatientname,Printed, Status from tbLabInstrumentResults where InstrumentID = '" & InstrumentID & "' and ResultDate between #" & Format(Now, "mm/dd/yyyy") & " 00:00# and #" & Format(Now, "mm/dd/yyyy") & " 23:59#")
    If rstTemp.RecordCount >= Val(GetSetting("Lab-I", "Others", "WARNINGCOUNT", "0")) Then frmWarning.Show 1
End If
If isDBconnected = True Then
    CodeLoc = "GETUNPROCESSED"
    Set rstTemp = myConnection.Execute("select count(resultdata) as totalrecord from labunprocessed where InstrumentID = '" & InstrumentID & "'")
    If Val(rstTemp.Fields("totalrecord").Value) > 0 Then
        DoEvents
        CodeLoc = "OPENPROCESSING"
        If Not frmProcess.Visible Then frmProcess.Show 1
    End If
    CodeLoc = "QUERYSEARCH"
    Set rstResult = myConnection.Execute(getSearchString)
    If rstResult.RecordCount > 0 Then
        CodeLoc = "QUERYRETURN"
        Unload frmSearchResult
        If lblTitle.Caption = "Search" Then frmSearchResult.wTrans = "QC" Else frmSearchResult.wTrans = "Result"
        chkResultDate.SetFocus
        frmSearchResult.Show
        isResultVisible = True
    Else
        CodeLoc = "QUERYNORECORD"
        Unload frmSearchResult
        isResultVisible = False
        MsgBox "No Record Found...", vbOKOnly + vbInformation, App.ProductName
        chkResultDate.SetFocus
    End If
Else
    CodeLoc = "NOSERVERCONNECTION"
    MsgBox "No Server Connection. Please Check Database Settings.", vbInformation + vbOKOnly, App.ProductName
    chkResultDate.SetFocus
End If
errFound:
    If Err.Number <> 0 Then
        MsgBox CodeLoc & vbCrLf & Err.Number & ":" & Err.Description, vbInformation + vbOKOnly, "ERROR"
    End If
Me.MousePointer = 0

End Sub

Private Sub dtEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub



Private Sub dtstart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub Form_Activate()
setFormColors Me
chkResultDate.BackColor = fraBody(0).BackColor
chkResultDate.ForeColor = lblText(0).ForeColor
SearchReset = False
End Sub

Private Sub Form_Load()
    cboStatus.Clear
    cboStatus.AddItem "All"
    cboStatus.AddItem "All-Today"
    cboStatus.AddItem "Unedited"
    cboStatus.AddItem "Unedited-Today"
    cboStatus.AddItem "Printed"
    cboStatus.AddItem "Saved"
    cboStatus.ListIndex = 0
    cmdClear_Click
    chkResultDate.Value = 1
    If GetSetting("HostForms", "Forms", "SearchQC", "0") = "1" Then
        lblTitle.Caption = "Search - Control"
        lblText(1).Caption = "UC NAME:"
        lblText(2).Caption = "LOT #:"
        SaveSetting "HostForms", "Forms", "SearchQC", "1"
    End If
'    fillPatient
End Sub

Public Function getSearchString(Optional sortOrder As Integer = 1, Optional sortAsc As Boolean = False) As String
Dim ctr As Integer
Dim tempStr As String
Dim strOrderby As String
Dim cntRecord As Integer
tempStr = ""
Select Case sortOrder
    Case 0: strOrderby = " order by Status " & IIf(sortAsc, " asc", "desc")
    Case 1: strOrderby = " order by ResultDate " & IIf(sortAsc, " asc", "desc")
    Case 2
     Case 3
        If DBServerType = "LOCAL DATABASE" Then
            If lblTitle.Caption = "Search - Control" Then
                strOrderby = " order by val(lotnum) " & IIf(sortAsc, " asc", "desc")
            Else
                strOrderby = " order by val(accessionnum) " & IIf(sortAsc, " asc", "desc")
            End If
        ElseIf DBServerType = "ORACLE" Then
            strOrderby = " order by trunc(Accessionnum) " & IIf(sortAsc, " asc", "desc")
        Else
            strOrderby = " order by [Sequence #] " & IIf(sortAsc, " asc", "desc")
        End If
End Select
If chkResultDate.Value = 1 Then
    If DBServerType = "LOCAL DATABASE" Then
        tempStr = "ResultDate between #" & Format(dtstart.Value, "mm/dd/yyyy HH:mm") & "# and #" & Format(dtEnd.Value, "mm/dd/yyyy HH:mm") & "#"
    ElseIf DBServerType = "ORACLE" Then
        tempStr = "TO_CHAR(ResultDate,'MM/DD/YYYY HH24:MI') between '" & Format(dtstart.Value, "mm/dd/yyyy 00:00") & "' and '" & Format(dtEnd.Value, "mm/dd/yyyy 23:59'")
    Else
        tempStr = "ResultDate between '" & Format(dtstart.Value, "mm/dd/yyyy 00:00") & "' and '" & Format(dtEnd.Value, "mm/dd/yyyy 23:59'")
    End If
End If

If lblTitle.Caption = "Search - Control" Then
    If Len(Trim(txtPatientName.Text)) > 0 Then tempStr = tempStr & " and UCNAME like '" & Replace(txtPatientName.Text, "'", "''") & "%'"
    If Len(Trim(txtSequenceNum.Text)) > 0 Then tempStr = tempStr & " and LOTNUM = '" & Replace(txtSequenceNum.Text, "'", "''") & "'"
Else
    If Len(Trim(txtPatientName.Text)) > 0 Then tempStr = tempStr & " and (PatientName like '%" & Replace(txtPatientName.Text, "'", "''") & "%' or NewPatientName like '" & Replace(txtPatientName.Text, "'", "''") & "%')"
    If Len(Trim(txtSequenceNum.Text)) > 0 Then tempStr = tempStr & " and Accessionnum = '%" & Replace(txtSequenceNum.Text, "'", "''") & "'" & " or AccessionnumOrg = '" & Replace(txtSequenceNum.Text, "'", "''") & "'"
    Select Case cboStatus.Text
        Case "Unedited", "Unedited-Today"
            tempStr = tempStr & " and status = '0' and Printed = False and (len(refnum) = 0 or isnull(refnum))"
        Case "Printed"
            tempStr = tempStr & " and Printed = True"
        Case "Saved"
            tempStr = tempStr & " and printed = FALSE and len(refnum) > 0"
    End Select
End If

If Mid(tempStr, 1, 5) = " and " Then tempStr = Mid(tempStr, 5)
If Len(Trim(tempStr)) > 0 Then tempStr = " and " & tempStr

cntRecord = Val(GetSetting("Lab-I", "Settings", "SearchCount", "300"))

If DBServerType = "LOCAL DATABASE" Then
    If lblTitle.Caption = "Search - Control" Then
        getSearchString = "Select Distinct Resultdate as [Result Date], UCNAME as [UC NAME], LOTNUM as [LOT #], val(LOTNUM) from tbLabInstrumentQC where InstrumentID = '" & InstrumentID & "'" & tempStr & strOrderby
    Else
        'getSearchString = "Select Distinct Resultdate as [Result Date], PatientName as [Patient Name], Accessionnum as [Sequence #], val(accessionnum), refnum, patientname, newpatientname from tbLabInstrumentResults where InstrumentID = '" & InstrumentID & "'" & tempStr & strOrderby
        getSearchString = "Select Distinct top " & cntRecord & " Resultdate as [Result Date], PatientName as [Patient Name], Accessionnum as [Sequence #],Accessionnum, refnum, patientname, newpatientname,Printed, Status, ResultID from tbLabInstrumentResults where InstrumentID = '" & InstrumentID & "'" & tempStr & strOrderby
    End If
ElseIf DBServerType = "ORACLE" Then
    getSearchString = "Select Distinct to_char(Resultdate,'yyyy-mm-dd HH24:mi:ss') ""Result Date"", PatientName ""Patient Name"", Accessionnum ""Sequence #"", resultdate from <schema>tbLabInstrumentResults where InstrumentID = '" & InstrumentID & "'" & tempStr & strOrderby
    getSearchString = Replace(getSearchString, "<schema>", DBSchema)
ElseIf DBServerType = "MSSQL SERVER" Then
    getSearchString = "Select Distinct Resultdate as [Result Date], PatientName as [Patient Name], Accessionnum as [Sequence #],Accessionnum, refnum, patientname, newpatientname,Printed, Status, ResultID  from tbLabInstrumentResults where InstrumentID = '" & InstrumentID & "'" & tempStr & strOrderby
ElseIf DBServerType = "MYSQL SERVER" Then
    getSearchString = "Select Distinct Resultdate ""Result Date"", PatientName ""Patient Name"", Accessionnum ""Sequence #"",Accessionnum, refnum, patientname, newpatientname,Printed, Status, ResultID from tbLabInstrumentResults where InstrumentID = '" & InstrumentID & "'" & tempStr & strOrderby
End If
End Function

Private Sub Form_Resize()
On Error Resume Next
    Me.Left = 1
    If isConnected Then
        Me.Height = (mdiMain.ScaleHeight / 2) - 10
        frmReceiveData.Height = mdiMain.ScaleHeight / 2
        Me.Top = frmReceiveData.Height
    Else
        Me.Top = 1
        Me.Height = mdiMain.ScaleHeight
    End If
    frmSearch.fraBody(0).Height = frmSearch.Height - 300
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mdiMain.tbrMain.Buttons("Search").Value = tbrPressed Then
        mdiMain.tbrMain.Buttons("Search").Value = tbrUnpressed
        mdiMain.mnuSearch.Checked = False
    End If
    If isResultVisible Then frmSearchResult.Hide
End Sub

Private Sub lblText_Click(Index As Integer)
If Index = 4 Then
    If chkResultDate.Value = 1 Then chkResultDate.Value = 0 Else chkResultDate.Value = 1
End If
End Sub

Private Sub lblTitle_DblClick()
'cmdClear_Click
'If lblTitle.Caption = "Search" Then
'    lblTitle.Caption = "Search - Control"
'    Label1.Caption = "UC NAME:"
'    Label2.Caption = "LOT #:"
'    SaveSetting "HostForms", "Forms", "SearchQC", "1"
'Else
'    lblTitle.Caption = "Search"
'    Label1.Caption = "Patient Name:"
'    Label2.Caption = "Sequence #:"
'    SaveSetting "HostForms", "Forms", "SearchQC", "0"
'End If
End Sub

Private Sub txtPatientName_Change()
  Dim Location As Long
'  lstPatientName.ListIndex = SendMessage(lstPatientName.hWnd, _
'        LB_FINDSTRING, -1, ByVal CStr(txtPatientName.Text))
'
  If lstPatientName.ListIndex <> -1 Then
    Location = txtPatientName.SelStart
    txtPatientName.Text = lstPatientName.Text
    txtPatientName.SelStart = Location
    txtPatientName.SelLength = Len(txtPatientName) - Location
  End If
End Sub


Private Sub txtPatientName_GotFocus()
    Highlight
End Sub

Private Sub txtPatientName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            keybd_event VK_TAB, 0, 0, 0
            KeyCode = 0
        Case 8
            If txtPatientName.SelLength <> 0 Then    ' ...and we have a string...
'              txtPatientName.Text = Mid$(txtPatientName.Text, 1, txtPatientName.SelStart - 1)
              KeyCode = 0
            End If
        Case vbKeyDelete
            If txtPatientName.SelLength <> 0 And _
               txtPatientName.SelStart <> 0 Then     ' Then clear all of it
               txtPatientName.Text = ""
               KeyCode = 0
            End If
    End Select
End Sub

Private Sub txtPatientName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSequenceNum_GotFocus()
    Highlight
End Sub

Private Sub txtSequenceNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub
