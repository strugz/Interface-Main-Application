VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchRequest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search Patient"
   ClientHeight    =   6636
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   11064
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6636
   ScaleWidth      =   11064
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H8000000E&
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   10380
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   45
      Width           =   630
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H8000000E&
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9675
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   45
      Width           =   705
   End
   Begin VB.CommandButton cmdCancelSearchHIS 
      Caption         =   "Cancel Search HIS"
      Enabled         =   0   'False
      Height          =   390
      Left            =   1650
      TabIndex        =   6
      Top             =   6495
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Timer tmrSearchExtDB 
      Enabled         =   0   'False
      Left            =   11430
      Top             =   6480
   End
   Begin VB.TextBox txtPatientID 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      MaxLength       =   100
      TabIndex        =   3
      Top             =   120
      Width           =   7695
   End
   Begin VB.Frame fraBody 
      Height          =   5745
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   660
      Width           =   10995
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   14040
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   3210
         Width           =   795
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   5430
         Left            =   90
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   10800
         _ExtentX        =   19050
         _ExtentY        =   9589
         _Version        =   393216
         BackColor       =   -2147483644
         Cols            =   5
         FixedCols       =   0
         RowHeightMin    =   300
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483644
         BackColorUnpopulated=   -2147483637
         GridColor       =   -2147483644
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   16777215
         AllowBigSelection=   0   'False
         FocusRect       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin MSComctlLib.ImageList imglstIcon 
      Left            =   10680
      Top             =   6480
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchRequest.frx":0000
            Key             =   "newrequest"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchRequest.frx":03FA
            Key             =   "checkin"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchRequest.frx":081E
            Key             =   "ongoing"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchRequest.frx":0C43
            Key             =   "ready"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchRequest.frx":1047
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchRequest.frx":144F
            Key             =   "new"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchRequest.frx":184B
            Key             =   "printed"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchRequest.frx":1C6A
            Key             =   "saved"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchRequest.frx":207D
            Key             =   "unapprove"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchRequest.frx":246F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Searching HIS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   6576
      Visible         =   0   'False
      Width           =   13392
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Name/ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   4
      Top             =   180
      Width           =   2025
   End
End
Attribute VB_Name = "frmSearchRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mSearchString As String
Private rsSearch As New ADODB.Recordset
Private rsSearchExtDB As New ADODB.Recordset
Private SRCH_RES_LIMIT As String
Private VIS_HID_SRCH As String
'Private strHospTblName As String
'Private strHospID As String, strHospPLName As String, strHospPFName As String, _
'        strHospPMName As String, strHospPDOB As String, strHospPGender As String, _
'        strHospPCivilStatus As String, strHospPAddress1 As String, strHospPAddress2 As String, _
'        strHospPHICID As String, strHospPEmail As String, strHospPhone As String
        
Private EnableExtDB As String
Private SearchInterval As String
Private ExtRSAutoSave As String
Private Externalmsg As String
Private ExternalmsgSrchLogsFileName As String
Private CtrQryExtDB As Integer
Private RowNumExtDB As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelSearchHIS_Click()
    delPrevExtDB
    If rsSearchExtDB.State = 1 Then
        rsSearchExtDB.Close
    End If
    tmrSearchExtDB.Enabled = False
    strSQLStatement = "EXEC sp_tbQryExternalDB_delete " & intWorkstation & ""
    Set rsSearch = myConnection.Execute(strSQLStatement)
    lblText(1).Visible = False
    lblText(1).Caption = "Searching HIS"
    cmdCancelSearchHIS.Visible = False
    cmdCancelSearchHIS.Enabled = False
    SearchLocalDB
End Sub

Private Sub cmdNew_Click()
    
    If EnableExtDB = "1" Then
        delPrevExtDB
        If rsSearchExtDB.State = 1 Then
            rsSearchExtDB.Close
        End If
        tmrSearchExtDB.Enabled = False
        strSQLStatement = "EXEC sp_tbQryExternalDB_delete " & intWorkstation & ""
        Set rsSearch = myConnection.Execute(strSQLStatement)
        lblText(1).Visible = False
        lblText(1).Caption = "Searching HIS"
        cmdCancelSearchHIS.Visible = False
        cmdCancelSearchHIS.Enabled = False
    End If

    With frmEditRequest
        .setBlnNewPatient = True
        .setUserName = strRequestByName
        .setUserID = strRequestBy
        .SetSectionID = globalSectionID
        If Len(Trim(txtPatientID.Text)) > 0 And Not IsNumeric(txtPatientID.Text) Then
            .setPatientName = Trim(txtPatientID.Text)
        End If
        .Show 1
    End With
    
End Sub

Private Sub cmdSearch_Click()

    If msfgResult.Visible Then
        msfgResult_DblClick
    Else
        If Len(Trim(txtPatientID.Text)) > 0 Then

            Me.MousePointer = vbHourglass
            
            With msfgResult
                .Visible = False
                .Enabled = False
                .Rows = 2
                .TextMatrix(1, 0) = ""
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
                .TextMatrix(1, 3) = ""
                .TextMatrix(1, 4) = ""
                .TextMatrix(1, 5) = ""
                .TextMatrix(1, 6) = ""
            End With
            
            If EnableExtDB = "1" Then
                If Not tmrSearchExtDB.Enabled Then
                    delPrevExtDB
                    InsertExtDBTblQry
                    'SendWSRequest "HOSTQUERY" & "|" & WorkStationName
                    tmrSearchExtDB.Interval = SearchInterval
                    lblText(1).Caption = "Searching HIS"
                    lblText(1).Visible = True
'                    cmdCancelSearchHIS.Visible = True
'                    cmdCancelSearchHIS.Enabled = True
                    tmrSearchExtDB.Enabled = True
                End If
                txtPatientID.SetFocus
                highLight
            Else
                SearchLocalDB
                If rsSearch.RecordCount = 0 Then
                    initGrid False
                    If MsgBox("Record Not Found, Do you want to add new record?", vbYesNo + vbQuestion, "Request") = vbYes Then
                        cmdNew_Click
                    End If
                    Me.Caption = "Search Patient"
                Else
                    Me.Caption = Me.Caption & " Result - [" & IIf(msfgResult.Rows = 2, "1", msfgResult.Rows - 2) & " Record(s) Found]"
                    With msfgResult
                        .Col = 0
                        .Row = 1
                        .ColSel = .Cols - 1
                        .TopRow = .Row
                        If Len(Trim(.TextMatrix(.Rows - 1, 1))) = 0 Then
                            .Rows = .Rows - 1
                        End If
                        .Visible = True
                        .Enabled = True
                    End With
                End If
            End If
        End If
    End If
    
    Me.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()
    txtPatientID.SetFocus
End Sub

Private Sub Form_Load()
    Dim intColCtr As Integer
    setFormColors Me
    
    If rstSettings.State = 0 Then
        getGlobalSettings
    End If

    SRCH_RES_LIMIT = FindSetting("SRCH_RES_LIMIT", "100")
    VIS_HID_SRCH = FindSetting("VIS_HID_SRCH", "0")
    EnableExtDB = GetSetting("HostForms_RE", "ExtDemoDB\Settings", "Enable", "0")
    Externalmsg = GetSetting("HostForms_RE", "ExtDemoDB\Settings", "Externalmsg", "Unable to connect to external database")
        
    cmdSearch.Picture = imglstIcon.ListImages(10).Picture
    cmdNew.Picture = imglstIcon.ListImages(1).Picture
    initGrid False
    If Len(Trim(mSearchString)) > 0 Then
        txtPatientID.Text = mSearchString
        cmdSearch_Click
        If rsSearch.RecordCount = 1 Then
            msfgResult.Row = 1
            msfgResult.ColSel = msfgResult.Cols - 1
            msfgResult_DblClick
            rsSearch.Close
        End If
    End If

    If VIS_HID_SRCH = 0 And EnableExtDB = 0 Then
        setSrchFrmDesign 1
    ElseIf VIS_HID_SRCH = 0 And EnableExtDB = 1 Then
        setSrchFrmDesign 2
        delPrevExtDB
'        OpenWSBroadcast
        SearchInterval = Val(GetSetting("HostForms_RE", "ExtDemoDB\Settings", "SearchInterval", "3")) * 1000
        ExtRSAutoSave = GetSetting("HostForms_RE", "ExtDemoDB\Settings", "AutoSave", "0")
    ElseIf VIS_HID_SRCH = 1 And EnableExtDB = 0 Then
        setSrchFrmDesign 3
    ElseIf VIS_HID_SRCH = 1 And EnableExtDB = 1 Then
        setSrchFrmDesign 4
        delPrevExtDB
'        OpenWSBroadcast
        SearchInterval = Val(GetSetting("HostForms_RE", "ExtDemoDB\Settings", "SearchInterval", "3")) * 1000
        ExtRSAutoSave = GetSetting("HostForms_RE", "ExtDemoDB\Settings", "AutoSave", "0")
    End If

    If EnableExtDB = "1" Then

        Me.Height = 7300
        CtrQryExtDB = 0
    End If
End Sub

Private Sub setSrchFrmDesign(DesignType As Integer)
    Select Case DesignType
        Case 1
            Me.Width = 11160
            Me.Height = 6990
            txtPatientID.Width = 7695
            cmdSearch.Left = 9675
            cmdNew.Left = 10380
            fraBody(0).Width = 10995
            msfgResult.Width = 10800
        Case 2
            Me.Width = 11160
            Me.Height = 7300
            txtPatientID.Width = 7695
            cmdSearch.Left = 9675
            cmdNew.Left = 10380
            fraBody(0).Width = 10995
            msfgResult.Width = 10800
        Case 3
            Me.Width = 13260
            Me.Height = 6990
            txtPatientID.Width = 9825
            cmdSearch.Left = 11805
            cmdNew.Left = 12510
            fraBody(0).Width = 13140
           msfgResult.Width = 12930
        Case 4
            Me.Width = 13260
            Me.Height = 7300
            txtPatientID.Width = 9825
            cmdSearch.Left = 11805
            cmdNew.Left = 12510
            fraBody(0).Width = 13140
            msfgResult.Width = 12930
    End Select
End Sub

Private Sub initGrid(blnEnable As Boolean)
    Dim intColCtr As Integer
    With msfgResult
            .Enabled = blnEnable
            If Not .Enabled Then
                .Rows = 2
                .Clear
            End If
            
            .Cols = 8
            
            If EnableExtDB = "1" Then
            
                .ColWidth(0) = 400
                .ColWidth(1) = 1600
                .ColWidth(2) = IIf(VIS_HID_SRCH = 1, 2100, 0)
                .ColWidth(3) = 2400
                .ColWidth(4) = 2400
                .ColWidth(5) = 2200
                .ColWidth(6) = 1500
                .ColWidth(7) = 0
                
                .TextArray(1) = "Lab Patient ID"
                .TextArray(2) = "Hospital Patient ID"
                .TextArray(3) = "Last Name"
                .TextArray(4) = "First Name"
                .TextArray(5) = "Middle Name"
                .TextArray(6) = "Date of Birth"
            Else
                .ColWidth(0) = 400
                .ColWidth(1) = 1600
                .ColWidth(2) = IIf(VIS_HID_SRCH = 1, 2100, 0)
                .ColWidth(3) = 2400
                .ColWidth(4) = 2400
                .ColWidth(5) = 2200
                .ColWidth(6) = 1500
                .ColWidth(7) = 0
                
                .TextArray(1) = "Lab Patient ID"
                .TextArray(2) = "Hospital Patient ID"
                .TextArray(3) = "Last Name"
                .TextArray(4) = "First Name"
                .TextArray(5) = "Middle Name"
                .TextArray(6) = "Date of Birth"
            End If

        End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CloseWSBroadcast
End Sub

Private Sub msfgResult_DblClick()

    With frmEditRequest
        If Val(msfgResult.TextMatrix(msfgResult.Row, 7)) = 0 Then
            .setPatientID = Trim(msfgResult.TextMatrix(msfgResult.Row, 1))
        Else
            .setPatientHospID = Trim(msfgResult.TextMatrix(msfgResult.Row, 2))
            If rsSearchExtDB.State = 1 Then
                rsSearchExtDB.Close
            End If
            tmrSearchExtDB.Enabled = False
            lblText(1).Visible = False
            lblText(1).Caption = "Searching HIS"
            cmdCancelSearchHIS.Visible = False
            cmdCancelSearchHIS.Enabled = False
        End If
        .SetSectionID = globalSectionID
        .setUserID = strRequestBy
        .setUserName = strRequestByName
        .setBlnNewPatient = False
        .Show 1
    End With
    
End Sub

Private Sub msfgResult_GotFocus()
    txtPatientID.SetFocus
End Sub

Private Sub msfgResult_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        msfgResult_DblClick
    End If
End Sub

Private Sub tmrSearchExtDB_Timer()
    
    If tmrSearchExtDB.Enabled = True Then
        
        tmrSearchExtDB.Enabled = False
        
        strSQLStatement = "EXEC sp_tbQryExternalDB_select " & Val(intWorkstation)
        
        DoEvents
        
        Set rsSearchExtDB = myConnection.Execute(strSQLStatement)
        
        If rsSearchExtDB.RecordCount > 0 Then
            If rsSearchExtDB.Fields("Status").Value = "1" Then
                strSQLStatement = "EXEC sp_PatientDemoTemp_select " & Val(intWorkstation)
                DoEvents
                Set rsSearchExtDB = myConnection.Execute(strSQLStatement)
                With msfgResult
                    For RowNumExtDB = 0 To rsSearchExtDB.RecordCount - 1
                        .Row = .Rows - 1
                        .Col = 0
                        .CellPictureAlignment = 3
                        Set .CellPicture = imglstIcon.ListImages(6).Picture
                        .Col = 1
                        .CellAlignment = 0
                        .TextMatrix(.Row, 1) = IIf(Not IsNull(rsSearchExtDB.Fields("PatientID").Value), rsSearchExtDB.Fields("PatientID").Value, "")
                        .Col = 2
                        .CellAlignment = 0
                        .TextMatrix(.Row, 2) = IIf(Not IsNull(rsSearchExtDB.Fields("HospitalID").Value), rsSearchExtDB.Fields("HospitalID").Value, "")
                        .Col = 3
                        .CellAlignment = 0
                        .TextMatrix(.Row, 3) = IIf(Not IsNull(rsSearchExtDB.Fields("LName").Value), rsSearchExtDB.Fields("LName").Value, "")
                        .Col = 4
                        .CellAlignment = 0
                        .TextMatrix(.Row, 4) = IIf(Not IsNull(rsSearchExtDB.Fields("FName").Value), rsSearchExtDB.Fields("FName").Value, "")
                        .Col = 5
                        .CellAlignment = 0
                        .TextMatrix(.Row, 5) = IIf(Not IsNull(rsSearchExtDB.Fields("MName").Value), rsSearchExtDB.Fields("MName").Value, "")
                        .Col = 6
                        .CellAlignment = 0
                        .TextMatrix(.Row, 6) = IIf(Not IsNull(rsSearchExtDB.Fields("DOB").Value), rsSearchExtDB.Fields("DOB").Value, "")
                        .TextMatrix(.Row, 7) = "1"
                        rsSearchExtDB.MoveNext
                        .Rows = .Rows + 1
                    Next
                    .Col = 0
                    .Row = 1
                    .ColSel = .Cols - 1
                End With
                    Me.Caption = Me.Caption & " Result - [" & IIf(msfgResult.Rows = 2, "1", msfgResult.Rows - 2) & " Record(s) Found In HIS] And "
                    txtPatientID.SetFocus
                    highLight
                    rsSearchExtDB.Close
                    lblText(1).Visible = False
                    lblText(1).Caption = "Searching HIS"
                    cmdCancelSearchHIS.Visible = False
                    cmdCancelSearchHIS.Enabled = False
                    SearchLocalDB
                    If rsSearch.RecordCount = 0 Then
                        msfgResult.Rows = msfgResult.Rows - 1
                    Else
                        Me.Caption = Me.Caption & "[" & rsSearch.RecordCount & " Record(s) Found In LIS]"
                    End If
                    msfgResult.Visible = True
            ElseIf rsSearchExtDB.Fields("Status").Value = "2" Then
                lblText(1).Caption = "No Record Found"
                SearchLocalDB
                If rsSearch.RecordCount = 0 Then
                    initGrid False
                    If MsgBox("Record Not Found, This record is not linked with the Hospital System, Result will not be transmitted to Hospital System" & _
                        " , Do you want to add new record?", vbYesNo + vbQuestion, "Request") = vbYes Then
                        cmdNew_Click
                    End If
                    Me.Caption = "Search Patient"
                End If
            ElseIf rsSearchExtDB.Fields("Status").Value = "3" Then
                lblText(1).Caption = Externalmsg
                MsgBox Externalmsg, , "LMS"
                SearchLocalDB
            ElseIf rsSearchExtDB.Fields("Status").Value = "0" Then
                If lblText(1).Caption = "Searching HIS..." Then
                    lblText(1).Caption = "Searching HIS"
                Else
                    lblText(1).Caption = lblText(1).Caption & "."
                End If
                tmrSearchExtDB.Enabled = True
            End If
        End If
       
    End If
End Sub

Private Sub txtPatientID_Change()
    msfgResult.Visible = False
    If EnableExtDB = "1" Then
        delPrevExtDB
        If rsSearchExtDB.State = 1 Then
            rsSearchExtDB.Close
        End If
        tmrSearchExtDB.Enabled = False
        strSQLStatement = "EXEC sp_tbQryExternalDB_delete " & intWorkstation & ""
        Set rsSearch = myConnection.Execute(strSQLStatement)
        lblText(1).Visible = False
        lblText(1).Caption = "Searching HIS"
        cmdCancelSearchHIS.Visible = False
        cmdCancelSearchHIS.Enabled = False
    End If
    Me.Caption = "Search Patient"
End Sub

Private Sub txtPatientID_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtPatientID_KeyDown(KeyCode As Integer, Shift As Integer)
    With msfgResult
        Select Case KeyCode
            Case vbKeyDown
                If .Row + 1 <> .Rows Then
                    .Row = .Row + 1
                    .ColSel = .Cols - 1
                Else
                    .Row = 1
                    .ColSel = .Cols - 1
                    .TopRow = .Row
                End If
                If Not .RowIsVisible(.Row) Then
                    .TopRow = .TopRow + 1
                End If
                KeyCode = 0
                highLight
            Case vbKeyUp
                If .Row <> 1 Then
                    .Row = .Row - 1
                    .ColSel = .Cols - 1
                Else
                    .Row = .Rows - 1
                    .ColSel = .Cols - 1
                    .TopRow = .Row
                End If
                If Not .RowIsVisible(.Row) Then
                    .TopRow = .TopRow - 1
                End If
                KeyCode = 0
                highLight
            Case vbKeyPageDown
                If .Row + 12 <= .Rows Then
                    .Row = .Row + 12
                    .ColSel = .Cols - 1
                End If
                If Not .RowIsVisible(.Row) Then
                    .TopRow = .TopRow + 12
                End If
                KeyCode = 0
                highLight
            Case vbKeyPageUp
                If .Row <> 1 Then
                    If .TopRow < 12 Then
                        .TopRow = 1
                        .ColSel = .Cols - 1
                    Else
                        .Row = .Row - 12
                        If .Row < 12 Then
                            .TopRow = 1
                        End If
                        .ColSel = .Cols - 1
                    End If
                End If
                If Not .RowIsVisible(.Row) Then
                    If .TopRow > 12 Then
                        .TopRow = .TopRow - 12
                    End If
                End If
                KeyCode = 0
                highLight
        End Select
    End With
End Sub

Private Sub txtPatientID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearch_Click
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Public Property Let wSearchString(ByVal vNewValue As String)
    mSearchString = vNewValue
End Property
 
Private Sub InsertExtDBTblQry()
    
    Dim rsInsertQry As New ADODB.Recordset
    
    strSQLStatement = "EXEC sp_tbQryExternalDB_insert " & intWorkstation & ",'" & txtPatientID.Text & "'"
    
    Set rsInsertQry = myConnection.Execute(strSQLStatement)
    
End Sub

Private Sub SearchLocalDB()
    
    Dim intRowCtr As Long

    strSQLStatement = "EXEC sp_PatientDemo_select '" & IIf(InStr(1, txtPatientID.Text, "'") > 0, Replace(txtPatientID.Text, "'", "''"), _
                    txtPatientID.Text) & "','" & SRCH_RES_LIMIT & "'"
                    
    Set rsSearch = myConnection.Execute(strSQLStatement)
    
    If rsSearch.RecordCount > 0 Then
        With msfgResult
            For intRowCtr = 0 To rsSearch.RecordCount - 1
                .Row = .Rows - 1
                .Col = 0
                .CellPictureAlignment = 3
                Set .CellPicture = imglstIcon.ListImages(8).Picture
                .Col = 1
                .CellAlignment = 0
                .TextMatrix(.Row, 1) = IIf(Not IsNull(rsSearch.Fields("PatientID").Value), rsSearch.Fields("PatientID").Value, "")
                .Col = 2
                .CellAlignment = 0
                .TextMatrix(.Row, 2) = IIf(Not IsNull(rsSearch.Fields("HospitalID").Value), rsSearch.Fields("HospitalID").Value, "")
                .Col = 3
                .CellAlignment = 0
                .TextMatrix(.Row, 3) = IIf(Not IsNull(rsSearch.Fields("LName").Value), rsSearch.Fields("LName").Value, "")
                .Col = 4
                .CellAlignment = 0
                .TextMatrix(.Row, 4) = IIf(Not IsNull(rsSearch.Fields("FName").Value), rsSearch.Fields("FName").Value, "")
                .Col = 5
                .CellAlignment = 0
                .TextMatrix(.Row, 5) = IIf(Not IsNull(rsSearch.Fields("MName").Value), rsSearch.Fields("MName").Value, "")
                .Col = 6
                .CellAlignment = 0
                .TextMatrix(.Row, 6) = IIf(Not IsNull(rsSearch.Fields("DOB").Value), rsSearch.Fields("DOB").Value, "")
                .TextMatrix(.Row, 7) = "0"
                rsSearch.MoveNext
                If .Rows <= rsSearch.RecordCount Then
                    .Rows = .Rows + 1
                End If
            Next
            .Visible = True
            .Col = 0
            .Row = 1
            .ColSel = .Cols - 1
        End With
    End If
     
End Sub

Private Sub AutoSaveDataFromExtDB()
    Dim rsAutoSave As New ADODB.Recordset
    Dim strToSelectExt As String
End Sub

Private Sub delPrevExtDB()
    Dim rsDelExtDB As New ADODB.Recordset
    Set rsDelExtDB = myConnection.Execute("EXEC sp_PatientDemoTemp_delete " & Val(intWorkstation))
End Sub

'Private Sub OpenWSBroadcast()
'    wsBroadcast.Protocol = sckUDPProtocol
'    wsBroadcast.LocalPort = 1007
'    wsBroadcast.RemotePort = 1007
'    wsBroadcast.RemoteHost = "255.255.255.255"
'    'WorkStationName = wsBroadcast.LocalHostName
'End Sub
'
'Private Sub CloseWSBroadcast()
'    wsBroadcast.RemoteHost = "255.255.255.255"
'    wsBroadcast.SendData wsBroadcast.LocalHostName
'End Sub
'
'Public Sub SendWSRequest(myData As String)
'    wsBroadcast.RemotePort = 1007
'    wsBroadcast.RemoteHost = "255.255.255.255"
'    wsBroadcast.SendData Blowfish.EncryptString(myData, myKey, True)
'End Sub

