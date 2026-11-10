VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmReceiveData 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Received Result Display Window"
   ClientHeight    =   4215
   ClientLeft      =   8310
   ClientTop       =   5055
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmHL7 
      Interval        =   6000
      Left            =   2790
      Top             =   240
   End
   Begin VB.Timer tmrWorklist 
      Interval        =   5000
      Left            =   2340
      Top             =   240
   End
   Begin VB.TextBox txtCommand 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   12
      Top             =   3420
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.Timer tmrStatus 
      Interval        =   1000
      Left            =   930
      Top             =   240
   End
   Begin VB.Timer tmrAutoPrint 
      Interval        =   5000
      Left            =   1410
      Top             =   240
   End
   Begin VB.Timer tmrAck 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1890
      Top             =   240
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   345
      Left            =   9120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   345
      Left            =   7740
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   345
      Left            =   6360
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   345
      Left            =   4980
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   3495
      Left            =   9120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "frmReceiveData.frx":0000
      Top             =   240
      Width           =   1305
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   3495
      Left            =   7740
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "frmReceiveData.frx":010D
      Top             =   240
      Width           =   1305
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   3495
      Left            =   6360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "frmReceiveData.frx":0112
      Top             =   240
      Width           =   1305
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   3495
      Left            =   4980
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "frmReceiveData.frx":0147
      Top             =   240
      Width           =   1305
   End
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6150
      TabIndex        =   1
      Top             =   60
      Width           =   1695
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receive Data"
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
         Left            =   150
         TabIndex        =   10
         Top             =   210
         Width           =   1290
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H00C0C0C0&
      Height          =   4095
      Index           =   0
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   960
      Top             =   780
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSCommLib.MSComm mscMain 
      Index           =   3
      Left            =   180
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm mscMain 
      Index           =   2
      Left            =   180
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm mscMain 
      Index           =   1
      Left            =   180
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm mscMain 
      Index           =   0
      Left            =   210
      Top             =   270
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin RichTextLib.RichTextBox rtbRcvDisplay 
      Height          =   3195
      Left            =   0
      TabIndex        =   11
      Top             =   150
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   5636
      _Version        =   393217
      BackColor       =   16777215
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmReceiveData.frx":0159
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblConsole 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Command:"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   3450
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Data Cable not Connected or DMS not Ready..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   90
      TabIndex        =   13
      Top             =   3900
      Width           =   3735
   End
End
Attribute VB_Name = "frmReceiveData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private AutoUnload As Boolean
Private Message(0 To 3) As String
Private TMessage(0 To 3) As String
Private AllMessage(0 To 3) As String
Private TotalMsgCount(0 To 3) As Integer
Private MsgCount(0 To 3) As Integer
Private ResultOtherInfo() As String
Private eConnection As New ADODB.Connection
Private rstExam2 As New ADODB.Recordset
Private lastError As String
Private HasHeader As Boolean
Private LastMessage As String
Private WaitCheckSum As Boolean
Private AckStr As String
Private ProcessNow As Boolean
Private WaitEnq As Boolean
Private WaitAck As Boolean
Private Nacked As Boolean
Private Syned As Boolean
Private SendMode As Boolean
Private ReceiveMode As Boolean
Private RequestSent As Boolean
Private ProcessRequest As Boolean
Private WLTimer As Integer
Private mRequestData As String
Private ProcessedResult As Boolean
Private WaitBlockCount As Boolean
Private ctrMsg As Integer
Private arrMsg() As String
Private dTrans As String
Private StopCapture As Boolean
Private blnLogTrans As Boolean
Private SaveResultFailed As Boolean  'set by SaveResult when the result could not be committed

Private MessageDef(0 To 3) As Object

Private cSOH(0 To 3)  As String
Private cSTX(0 To 3)  As String
Private cETX(0 To 3)  As String
Private cEOT(0 To 3)  As String
Private cENQ(0 To 3)  As String
Private cACK(0 To 3)  As String
Private cNAK(0 To 3)  As String
Private cSYN(0 To 3)  As String
 
Public Sub showConsole()
    If txtCommand.Visible = False Then
        rtbRcvDisplay.Height = Me.ScaleHeight - (Me.fraTitle.Height + 800)
        txtCommand.Top = rtbRcvDisplay.Height + 200
        lblConsole.Top = rtbRcvDisplay.Height + 200
        lblConsole.Visible = True
        txtCommand.Visible = True
        txtCommand.SetFocus
    Else
        rtbRcvDisplay.Height = Me.ScaleHeight - (Me.fraTitle.Height + 400)
        lblConsole.Visible = False
        txtCommand.Visible = False
        rtbRcvDisplay.SetFocus
    End If
    txtCommand.Text = ""
End Sub

Private Sub Command1_Click()
HandleInput Text1.Text
End Sub

Private Sub Command2_Click()
HandleInput Text2.Text
End Sub

Private Sub Command3_Click()
HandleInput Text3.Text
End Sub

Private Sub Command4_Click()
HandleInput Text4.Text
End Sub

Private Sub Form_Activate()
setFormColors Me
End Sub

Private Sub Form_Load()
    If mdiMain.mnuSearch.Checked Then
        Me.Height = mdiMain.ScaleHeight / 2
    Else
        Me.Height = mdiMain.ScaleHeight
    End If
    Message(0) = ""
    Message(1) = ""
    Message(2) = ""
    Message(3) = ""
    MsgCount(0) = 0
    MsgCount(1) = 0
    MsgCount(2) = 0
    MsgCount(3) = 0
    TotalMsgCount(0) = 0
    TotalMsgCount(1) = 0
    TotalMsgCount(2) = 0
    TotalMsgCount(3) = 0
    HasHeader = False
    ProcessNow = False
    mdiMain.mnuConnect.Enabled = False
    mdiMain.mnuDisconnect.Enabled = True
    mdiMain.mnuLabExam.Enabled = False
    AutoUnload = False
    Me.Top = 0
    Me.Left = 1
    isConnected = True
    lblStatus.Caption = "Data Cable not Connected or DMS not Ready..."
    lblTitle.Caption = "Receive Data - [Disconnected]"
    OpenLocalConnection
    If isDBconnected Then
        If DBServerType = "LOCAL DATABASE" Then
            Set rstExam2 = eConnection.Execute("select * from tbcolabexam where Active = True and InstrumentID = '" & InstrumentID & "' and HisExamCode <> '0' and HisResultCode <> '0'")
        ElseIf DBServerType = "MYSQL SERVER" Then
            Set rstExam2 = eConnection.Execute("select * from tbcolabexam where Active = True and InstrumentID = '" & InstrumentID & "' and HisExamCode <> '0' and HisResultCode <> '0'")
        End If
    End If
    checkComStatus
    SetComm
    WaitCheckSum = False
    AckStr = ""
    WaitEnq = False
    WaitAck = False
    Nacked = False
    Syned = False
    SendMode = False
    ReceiveMode = False
    ProcessRequest = False
    WLTimer = 0
    mRequestData = ""
    dTrans = "N"
'    If Val(GetSetting("Lab-I", "Settings", "Autoprint", "0")) = 1 Then
'        tmrAutoProcess.Interval = 5000
'    Else
'        tmrAutoProcess.Interval = 0
'    End If
    If Val(GetSetting("ComSet", "Settings", "LogTrans", "0")) = 1 Then blnLogTrans = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Me.fraBody(0).Height = Me.ScaleHeight - Me.fraTitle.Height + 160
    'rtbRcvDisplay.Height = Me.ScaleHeight - (Me.fraTitle.Height + 400)
    lblStatus.Top = fraBody(0).Height - 300
    If txtCommand.Visible = True Then
        rtbRcvDisplay.Height = Me.ScaleHeight - (Me.fraTitle.Height + 800)
        txtCommand.Top = rtbRcvDisplay.Height + 200
        lblConsole.Top = rtbRcvDisplay.Height + 200
    Else
        rtbRcvDisplay.Height = Me.ScaleHeight - (Me.fraTitle.Height + 400)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If AutoUnload = False Then
    If MsgBox("Are you sure you want to disconnect the host computer?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
        If AutoUnload = False Then
            If Val(GetSetting("ComSet", "Control 1", "Enabled", "0")) = 1 Then mscMain(0).PortOpen = False
            If Val(GetSetting("ComSet", "Control 2", "Enabled", "0")) = 1 Then mscMain(1).PortOpen = False
            If Val(GetSetting("ComSet", "Control 3", "Enabled", "0")) = 1 Then mscMain(2).PortOpen = False
            If Val(GetSetting("ComSet", "Control 4", "Enabled", "0")) = 1 Then mscMain(3).PortOpen = False
        End If
        isConnected = False
        If mdiMain.mnuSearch.Checked Then
            frmSearch.Top = 1
            frmSearch.Height = mdiMain.ScaleHeight  'mdiMain.Height - 800
        End If
        Set eConnection = Nothing
        Set rstExam2 = Nothing
        
        mdiMain.mnuConnect.Enabled = True
        mdiMain.mnuDisconnect.Enabled = False
        mdiMain.mnuLabExam.Enabled = True
        If isConnected = False Then setButtons 1, 0, 1, 1, 1
    Else
        If mdiMain.mnuSearch.Checked Then frmSearch.Show
        If isResultVisible Then frmSearchResult.Show
        Cancel = 1
    End If
Else
End If
End Sub

Private Sub SetComm()
Dim ctr As Integer
Dim tempstr As String
Dim tempStr2 As String
On Error GoTo errFound
For ctr = 1 To 4
    tempStr2 = "Control " & ctr
    mscMain(ctr - 1).CommPort = Replace(GetSetting("ComSet", tempStr2, "ComPort", "COM1"), "COM", "")
    tempstr = Trim(GetSetting("ComSet", tempStr2, "BaudRate", "9600")) & "," & Trim(Mid(GetSetting("ComSet", tempStr2, "Parity", "NONE"), 1, 1)) & "," & Trim(GetSetting("ComSet", tempStr2, "DataBits", "8")) & "," & Trim(GetSetting("ComSet", tempStr2, "StopBits", "1"))
    mscMain(ctr - 1).Settings = tempstr
    mscMain(ctr - 1).Handshaking = Val(Mid(GetSetting("ComSet", tempStr2, "HandShake", "0-ComNone"), 1, 1))
    mscMain(ctr - 1).InBufferSize = Val(GetSetting("ComSet", tempStr2, "InbufferSize", "1024"))
    mscMain(ctr - 1).OutBufferSize = Val(GetSetting("ComSet", tempStr2, "OutbufferSize", "512"))
    mscMain(ctr - 1).RThreshold = Val(GetSetting("ComSet", tempStr2, "RTresHold", "1"))
    mscMain(ctr - 1).SThreshold = Val(GetSetting("ComSet", tempStr2, "STresHold", "1"))
    mscMain(ctr - 1).InputLen = Val(GetSetting("ComSet", tempStr2, "InputLen", "1"))
    mscMain(ctr - 1).ParityReplace = GetSetting("ComSet", tempStr2, "ParityReplace", "?")
    mscMain(ctr - 1).EOFEnable = Val(GetSetting("ComSet", tempStr2, "EOFEnable", "0"))
    mscMain(ctr - 1).NullDiscard = Val(GetSetting("ComSet", tempStr2, "NullDiscard", "1"))
    mscMain(ctr - 1).RTSEnable = Val(GetSetting("ComSet", tempStr2, "RTSEnable", "1"))
    mscMain(ctr - 1).DTREnable = Val(GetSetting("ComSet", tempStr2, "DTREnable", "1"))
    If Val(GetSetting("ComSet", tempStr2, "Enabled", "0")) = 1 Then
        mscMain(ctr - 1).PortOpen = True
    End If
Next ctr
    If mdiMain.mnuSearch.Checked Then
        Me.Height = mdiMain.ScaleHeight / 2
        frmSearch.Top = Me.Height
        frmSearch.Height = mdiMain.ScaleHeight / 2 '- frmReceiveData.Height  'mdiMain.Height - 800 - frmReceiveData.Height
    End If
errFound:
    If Err.Number = 8005 Then
        MsgBox "Another program is using the selected Telephony device." & vbCrLf & "Try again after the other program completes.", vbInformation + vbOKOnly, App.ProductName
        AutoUnload = True
        mdiMain.mnuConnect.Enabled = True
        mdiMain.mnuDisconnect.Enabled = False
        isConnected = False
        Unload Me
    ElseIf Err.Number = 8002 Or Err.Number = 13 Then
        MsgBox "Invalid Port Number." & vbCrLf & "Please Check the port number.", vbInformation + vbOKOnly, App.ProductName
        AutoUnload = True
        mdiMain.mnuConnect.Enabled = True
        mdiMain.mnuDisconnect.Enabled = False
        isConnected = False
        Unload Me
    ElseIf Err.Number = 0 Then
        displayresults "~*~ Connected at " & Now & "~*~"
        isConnected = True
        'DisplayResults "COMMUNICATION CONNECTION: OK"
        If isDBconnected = False Then
            displayresults "DATABASE ERRROR: NO SERVER CONNECTION"
        End If
    Else
        MsgBox Err.Number & ": " & Err.Description, vbExclamation + vbOKOnly, App.ProductName
        AutoUnload = True
        mdiMain.mnuConnect.Enabled = True
        mdiMain.mnuDisconnect.Enabled = False
        isConnected = False
        Unload Me
    End If
End Sub

Private Sub lblTitle_Change()
    Me.Caption = lblTitle.Caption
End Sub



Private Sub mscMain_OnComm(Index As Integer)
Dim InBuff(0 To 4) As String
Dim tempstr As String
Dim tempStr2 As String
Dim rstTemp As ADODB.Recordset
Select Case mscMain(Index).CommEvent
    Case comEvDSR, comEvCD, comEvCTS, comEventBreak
        checkComStatus
    Case comEventCDTO
        lblStatus.Caption = "CD (RLSD) Timeout."
        displayresults lblStatus.Caption
    Case comEventCTSTO
        lblStatus.Caption = "CTS Timeout."
        displayresults lblStatus.Caption
    Case comEventDSRTO
        lblStatus.Caption = "DSR Timeout."
        displayresults lblStatus.Caption
    Case comEventFrame
        If lastError <> "Framing Error." Then
            lblStatus.Caption = "Framing Error."
        End If
        displayresults lblStatus.Caption
    Case comEventOverrun
        lblStatus.Caption = "Data Lost."
        displayresults lblStatus.Caption
    Case comEventRxOver
        If lastError <> "Receive buffer overflow." Then
            lblStatus.Caption = "Receive buffer overflow."
        End If
        lastError = lblStatus.Caption
        InBuff(Index) = InBuff(Index) + mscMain(Index).Input
        mscMain(Index).PortOpen = False
        mscMain(Index).PortOpen = True
        mscMain(Index).RTSEnable = True
        mscMain(Index).DTREnable = True
        displayresults lblStatus.Caption
    Case comEventRxParity
        lblStatus.Caption = "Parity Error."
        displayresults lblStatus.Caption
    Case comEventTxFull
        lblStatus.Caption = "Transmit buffer full."
        displayresults lblStatus.Caption
    Case comEventDCB
        lblStatus.Caption = "Unexpected error retrieving DCB"
        displayresults lblStatus.Caption
    Case comEvReceive
        If Val(GetSetting("ComSet", "Control 1", "FullHandshake", "0")) = 0 Then
            If LoopBack Then
                lastError = ""
                InBuff(Index) = mscMain(Index).Input
                displayresults InBuff(Index)
                mscMain(Index).Output = InBuff(Index)
                InBuff(Index) = ""
            Else
                Screen.MousePointer = 11
                InBuff(Index) = mscMain(Index).Input
                lastError = ""
                If InBuff(Index) = Chr(5) Then
                    sendAck Index
                Else
                '------------- Remove: Gem3000, Gem3500, Gem4000 abnormal communication
                    If MessageFormat = "ASTM1394" Then
                        If InBuff(Index) = Chr(3) Then
                            sendAck Index
                            StopCapture = False
                        ElseIf InBuff(Index) = Chr(23) Then
                            sendAck Index
                            StopCapture = True
                        ElseIf InBuff(Index) = Chr(2) Then
                            sendAck Index
                            StopCapture = False
                        ElseIf InBuff(Index) = Chr(5) Or InBuff(Index) = Chr(4) Then
                            StopCapture = False
                            sendAck Index
                        End If
                    End If
'                If MsgDef.InstrumentID = "IQ200" Then
'                    If InStr(1, InBuff(Index), vbCrLf) > 0 Then
'                        sendAck Index
'                    End If
'                End If
                If Not StopCapture Then
                    HandleInput InBuff(Index), Index
                End If
            End If
                InBuff(Index) = ""
                Screen.MousePointer = 0
            End If
        Else
            If LoopBack Then
                lastError = ""
                InBuff(Index) = mscMain(Index).Input
                displayresults InBuff(Index)
                mscMain(Index).Output = InBuff(Index)
                InBuff(Index) = ""
            ElseIf Val(Mid(PluginVersion, 1, InStr(1, PluginVersion, ".") - 1)) >= 2 And Not LoopBack Then
                InBuff(Index) = mscMain(Index).Input
                If InStr(1, InBuff(Index), Chr(4)) = False Then
                    InBuff(Index) = InBuff(Index) & mscMain(Index).Input
                End If

                If blnLogTrans Then
                    LogComTrans "<DMS> " & InBuff(Index), Index
                End If
                tempstr = MsgDef.HandleInput(InBuff(Index))
                If Len(tempstr) > 0 Then

                    If MsgDef.HasQuery Then
                        If MsgDef.CurrentSampleIDCount = 1 Then
                            tempStr2 = MsgDef.CurrentSampleID
'                            MsgDef(Index).SetSampleIDValues = GetSampleRequest(tempStr2, HostIDArr(Index), InstID(Index))
                            SendtoCom 2, tempstr, Index
                            tempstr = ""
                        Else
                            SendtoCom 2, tempstr, Index
                            tempstr = ""
                        End If
                    Else
                        SendtoCom 2, tempstr, Index
                        tempstr = ""
                    End If
                End If
                If MsgDef.HasResult Then
                    processData MsgDef.RawData, Index
'                    MsgDef.HasResult = False
                    MsgDef.RawData True
'                    StrHostQuery(Index) = ""
                    AllMessage(Index) = ""
                End If
            Else
                Screen.MousePointer = 11
                lastError = ""
                InBuff(Index) = mscMain(Index).Input
                If MessageFormat = "STKS" Then
                    If ProcessedResult And InBuff(Index) = Chr(22) Then
                        ProcessedResult = False
                        ReceiveMode = False
                        SendMode = False
                        InBuff(Index) = ""
                        sendAck Index
                    End If
                    If InBuff(Index) = Chr(22) Then
                        mscMain(Index).Output = Chr(22)
                        ReceiveMode = True
                        SendMode = False
                        ProcessedResult = False
                        WaitBlockCount = True
                    End If
                    If ReceiveMode Then
                        Screen.MousePointer = 11
                        If InBuff(Index) = Chr(5) Or InBuff(Index) = Chr(3) Then
                            sendAck Index
                            HandleInput InBuff(Index), Index
                        ElseIf InBuff(Index) = Chr(22) Then
                            sendAck Index
                            HandleInput InBuff(Index), Index
                        ElseIf WaitBlockCount Then
                            WaitBlockCount = False
                            HandleInput InBuff(Index), Index
                        Else
                            HandleInput InBuff(Index), Index
                        End If
                        InBuff(Index) = ""
                        Screen.MousePointer = 0
                    End If
                ElseIf MessageFormat = "ASTM" Then
                    Screen.MousePointer = 11
                    lastError = ""
'                    Debug.Print InBuff(Index)
'                    InBuff(Index) = mscMain(Index).Input
                    If InBuff(Index) = Chr(5) Then
                        sendAck Index
                        SendMode = False
                        ReceiveMode = True
                        HandleInput InBuff(Index), Index
                    ElseIf InBuff(Index) = Chr(21) Then
                        SendMode = False
                        ReceiveMode = False
                        WaitAck = False
                        Syned = True
                        AllMessage(Index) = ""
                        TMessage(Index) = ""
                    Else
                        If ReceiveMode Then HandleInput InBuff(Index), Index
                        If ProcessedResult Then
                            ReceiveMode = False
                            ProcessedResult = False
                        End If
                    End If
'                    InBuff(Index) = ""
                    Screen.MousePointer = 0
                End If
                If SendMode Then
                    If MessageFormat = "STKS" Then
                        If WaitEnq Then
                            If InBuff(Index) = Chr(5) Then
                                SendMode = True
                                Set rstTemp = myConnection.Execute("Select top 1 * from tblabrequest where messagestatus = '1'")
                                dFilename = rstTemp.Fields("messageid").Value
                                mRequestData = ProcessHL7Request(rstTemp.Fields("messagedata").Value, IIf(IsNull(rstTemp.Fields("MessageRack").Value), "", rstTemp.Fields("MessageRack").Value), IIf(IsNull(rstTemp.Fields("MessageCassPos").Value), "", rstTemp.Fields("MessageCassPos").Value))
                                mscMain(Index).Output = Left(mRequestData, 2)
                                WaitEnq = False
                                WaitAck = True
                            ElseIf InBuff(Index) = Chr(22) Then
                                Syned = True
                                WaitEnq = False
                                SendMode = False
                                dFilename = ""
                            End If
                        ElseIf WaitAck Then
                            If Not RequestSent Then
                                mscMain(Index).Output = Mid(mRequestData, 3)
                                Debug.Print Mid(mRequestData, 3)
                                RequestSent = True
                                WaitAck = True
                            Else
                                mscMain(Index).Output = Chr(5)
                                If Len(Trim(dFilename)) > 0 Then myConnection.Execute ("Update tblabrequest set messagestatus = '2' where messageid = " & dFilename)
                                dFilename = ""
                                RequestSent = False
                                WaitAck = False
                                WaitEnq = False
                                Syned = False
                                SendMode = False
                            End If
                        Else
                        End If
                    ElseIf MessageFormat = "ASTM" Then
                        If WaitAck Then
                            If InBuff(Index) = Chr(6) Then
                                If Not ProcessRequest Then
                                    Set rstTemp = myConnection.Execute("Select top 1 * from tblabrequest where messagestatus = '1'")
                                    dFilename = rstTemp.Fields("messageid").Value
                                    dTrans = rstTemp.Fields("MessageTrans").Value
                                    mRequestData = ProcessHL7Request(rstTemp.Fields("messagedata").Value, rstTemp.Fields("MessageRack").Value, rstTemp.Fields("MessageCassPos").Value & "", rstTemp.Fields("MessageTrans").Value) & ""
                                    arrMsg = Split(mRequestData, vbCrLf) & ""
                                    ctrMsg = 0
                                    ProcessRequest = True
                                End If
                                If ctrMsg < UBound(arrMsg) Then
                                    mscMain(Index).Output = arrMsg(ctrMsg) & vbCrLf
                                    ctrMsg = ctrMsg + 1
                                ElseIf ctrMsg = UBound(arrMsg) Then
                                    mscMain(Index).Output = arrMsg(ctrMsg)
                                    ctrMsg = ctrMsg + 1
                                    If Trim(dTrans) = "C" Then
                                        myConnection.Execute ("Update tblabrequest set messagestatus = '0' where messageid = " & dFilename)
                                        dTrans = "N"
                                    Else
                                        If Len(Trim(dFilename)) > 0 Then myConnection.Execute ("Update tblabrequest set messagestatus = '2' where messageid = " & dFilename)
                                        dTrans = "N"
                                    End If
                                    ctrMsg = 0
                                    ProcessRequest = False
                                    SendMode = False
                                    ReceiveMode = False
                                    WaitAck = False
                                Else
                                    ctrMsg = 0
                                    ProcessRequest = False
                                    SendMode = False
                                    ReceiveMode = False
                                    WaitAck = False
                                End If
                            End If
                        End If
                    End If
                End If
                Screen.MousePointer = 0
            End If
        End If
    Case comEvSend
        lblStatus.Caption = "Acknowleged..."
    Case Else
        lblStatus.Caption = "Unknown Event: " & mscMain(Index).CommEvent
        displayresults lblStatus.Caption
End Select
End Sub

Private Sub HandleInput(mess As String, Optional portIndex As Integer = 0)
Dim numstart(0 To 3) As Integer
Dim numend(0 To 3) As Integer
Dim dateStart(0 To 3) As Date
On Error Resume Next
Message(portIndex) = Message(portIndex) & mess
If MsgDef.InstrumentID = "IQ200" Then
    lblStatus.Caption = "Receiving Data..."
    If InStr(1, Message(portIndex), Chr(2)) > 0 Then
        If InStr(1, Message(portIndex), Chr(3)) > 0 Then
            TMessage(portIndex) = Message(portIndex)
            AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
            processData AllMessage(portIndex)
            AllMessage(portIndex) = ""
            TMessage(portIndex) = ""
            Message(portIndex) = ""
        End If
    End If
'Else
'    lblStatus.Caption = "Receiving Data..."
'    TMessage(portIndex) = Message(portIndex)
'    AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
'    processData AllMessage(portIndex)
'    AllMessage(portIndex) = ""
'    TMessage(portIndex) = ""
'    Message(portIndex) = ""
End If
If InStr(1, Message(portIndex), Chr(2)) > 0 Then
    lblStatus.Caption = "Receiving Data..."
'    If MessageFormat = "ASTM1394" Then sendAck portIndex
    If InStr(1, Message(portIndex), Chr(3)) > 0 Then
        Select Case MessageFormat
            Case "STKS"
                If InStr(1, Message(portIndex), Chr(2) & "01") > 0 Then
                    TMessage(portIndex) = PartMessage(Message(portIndex))
                    TotalMsgCount(portIndex) = Val(Mid(Message(portIndex), InStr(1, Message(portIndex), Chr(2) & "01") - 2, 2))
                    MsgCount(portIndex) = 1
                    AllMessage(portIndex) = TMessage(portIndex)
                    Message(portIndex) = Mid(Message(portIndex), InStr(1, Message(portIndex), Chr(3)) + 1)
                Else
                    'MsgCount(portIndex) = Val(Mid(Message(portIndex), InStr(1, Message(portIndex), Chr(2)) + 1, 2))
                    If MsgCount(portIndex) > 0 Then
                        MsgCount(portIndex) = Val(Mid(Message(portIndex), InStr(1, Message(portIndex), Chr(2)) + 1, 2))
                        TMessage(portIndex) = PartMessage(Message(portIndex))
                        rtbRcvDisplay.SelStart = Len(rtbRcvDisplay.Text)
                        AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
                        If TotalMsgCount(portIndex) = MsgCount(portIndex) Then
'                            mscMain.DTREnable = False
'                            mscMain.RTSEnable = False
                            processData AllMessage(portIndex)
                            AllMessage(portIndex) = ""
                            TMessage(portIndex) = ""
                            MsgCount(portIndex) = 1
                            TotalMsgCount(portIndex) = 1
                            lblStatus.Caption = "Ready to receive data..."
                            ProcessedResult = True
'                            mscMain.RTSEnable = True
'                            mscMain.DTREnable = True
                        End If
                        Message(portIndex) = Mid(Message(portIndex), InStr(1, Message(portIndex), Chr(3)) + 1)
                    End If
                End If
            Case "ASTM"
                If InStr(Len(Message(portIndex)) - 1, Message(portIndex), vbLf) > 0 Then
                    If getASTMmsgType(Message(portIndex)) = "H" Then
                        If Len(Trim(AllMessage(portIndex))) = 0 Then
                            TMessage(portIndex) = ""
                            AllMessage(portIndex) = Message(portIndex)
                            Message(portIndex) = ""
                            Do While DateDiff("s", dateStart(portIndex), Now) <= 2
                                'wait 3 sec.
                            Loop
                            sendAck portIndex
                        Else
                            TMessage(portIndex) = Message(portIndex)
                            AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
                            processData AllMessage(portIndex)
                            AllMessage(portIndex) = ""
                            TMessage(portIndex) = ""
                            Message(portIndex) = ""
                            Do While DateDiff("s", dateStart(portIndex), Now) <= 2
                                'wait 3 sec.
                            Loop
                            sendAck portIndex
                        End If
                    ElseIf getASTMmsgType(Message(portIndex)) = "L" Then
                        TMessage(portIndex) = Message(portIndex)
                        AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
                        processData AllMessage(portIndex)
                        AllMessage(portIndex) = ""
                        TMessage(portIndex) = ""
                        Message(portIndex) = ""
                        Do While DateDiff("s", dateStart(portIndex), Now) <= 2
                            'wait 3 sec.
                        Loop
                        sendAck portIndex
                        ProcessedResult = True
                    Else
                        TMessage(portIndex) = Message(portIndex)
                        AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
                        Message(portIndex) = ""
                        Do While DateDiff("s", dateStart(portIndex), Now) <= 2
                            'wait 3 sec.
                        Loop
                        sendAck portIndex
                    End If
                End If
            Case "ASTM2"
                If InStr(Len(Message(portIndex)) - 1, Message(portIndex), vbLf) > 0 Then
                    If getASTMmsgType(Message(portIndex)) = "H" Then
                        If Len(Trim(AllMessage(portIndex))) = 0 Then
                            TMessage(portIndex) = ""
                            AllMessage(portIndex) = Message(portIndex)
                            Message(portIndex) = ""
                            Do While DateDiff("s", dateStart(portIndex), Now) <= 2
                                'wait 3 sec.
                            Loop
                            sendAck portIndex
                        Else
                            TMessage(portIndex) = Message(portIndex)
                            AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
                            processData AllMessage(portIndex)
                            AllMessage(portIndex) = ""
                            TMessage(portIndex) = ""
                            Message(portIndex) = ""
                            Do While DateDiff("s", dateStart(portIndex), Now) <= 2
                                'wait 3 sec.
                            Loop
                            sendAck portIndex
                        End If
                    ElseIf getASTMmsgType(Message(portIndex)) = "L" Then
                        TMessage(portIndex) = Message(portIndex)
                        AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
                        processData AllMessage(portIndex)
                        AllMessage(portIndex) = ""
                        TMessage(portIndex) = ""
                        Message(portIndex) = ""
                        Do While DateDiff("s", dateStart(portIndex), Now) <= 2
                            'wait 3 sec.
                        Loop
                        sendAck portIndex
                    Else
                        TMessage(portIndex) = Message(portIndex)
                        AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
                        Message(portIndex) = ""
                        Do While DateDiff("s", dateStart(portIndex), Now) <= 2
                            'wait 3 sec.
                        Loop
                        sendAck portIndex
                    End If
                End If
            Case "ILABBG"
                TMessage(portIndex) = Message(portIndex)
                AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
                If InStr(1, AllMessage(portIndex), Chr(2) & "06") Then
                    processData AllMessage(portIndex)
                    AllMessage(portIndex) = ""
                End If
                TMessage(portIndex) = ""
                Message(portIndex) = ""
            Case "ASYNC"
                TMessage(portIndex) = Message(portIndex)
                Select Case MsgDef.Flags(2, TMessage(portIndex))
                    Case "RS"
                        AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
                        If InStr(1, AllMessage(portIndex), Chr(3)) Then
                            processData AllMessage(portIndex)
                        End If
                    Case "DTE"
                        processData AllMessage(portIndex)
                        AllMessage(portIndex) = ""
                    Case ""
                        Debug.Print TMessage(portIndex)
                    Case Else
                        Debug.Print MsgDef.Flags(2, TMessage(portIndex))
                End Select
                TMessage(portIndex) = ""
                Message(portIndex) = ""
            Case "STXETX"
                lblStatus.Caption = "Receiving Data..."
                TMessage(portIndex) = Message(portIndex)
                AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
                processData AllMessage(portIndex)
                AllMessage(portIndex) = ""
                TMessage(portIndex) = ""
                Message(portIndex) = ""
            Case "ASTM1394"
                TMessage(portIndex) = Message(portIndex)
                AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
                processData AllMessage(portIndex)
                AllMessage(portIndex) = ""
                TMessage(portIndex) = ""
                Message(portIndex) = ""
        End Select
    End If
End If
If MessageFormat = "ACTDIFF" Then
    lblStatus.Caption = "Receiving Data..."
    If InStr(1, Message(portIndex), vbCrLf) > 0 Then
        If getAcTmsgType(Message(portIndex)) = "H" Then
            If Len(Trim(AllMessage(portIndex))) = 0 Then
                TMessage(portIndex) = ""
                AllMessage(portIndex) = Message(portIndex)
                Message(portIndex) = ""
                sendAck portIndex
            Else
                TMessage(portIndex) = Message(portIndex)
                AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
                processData AllMessage(portIndex)
                lblStatus.Caption = "Ready to receive data..."
                AllMessage(portIndex) = ""
                TMessage(portIndex) = ""
                Message(portIndex) = ""
                sendAck portIndex
            End If
        ElseIf getAcTmsgType(Message(portIndex)) = "L" Then
            TMessage(portIndex) = Message(portIndex)
            AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
            processData AllMessage(portIndex)
            lblStatus.Caption = "Ready to receive data..."
            AllMessage(portIndex) = ""
            TMessage(portIndex) = ""
            Message(portIndex) = ""
            sendAck portIndex
        Else
            TMessage(portIndex) = Message(portIndex)
            AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
            Message(portIndex) = ""
            sendAck portIndex
        End If
    End If
End If
If MessageFormat = "CXSYS" Or MessageFormat = "LXSYS" Then
    lblStatus.Caption = "Receiving Data..."
    If InStr(1, Message(portIndex), "[") > 0 Then
        If InStr(1, Message(portIndex), "]") > 0 Then
            Select Case getCXSYSmsgType(Message(portIndex))
                Case "H"
                    AllMessage(portIndex) = ""
                    TMessage(portIndex) = Mid(Message(portIndex), InStr(1, Message(portIndex), "["), InStr(1, Message(portIndex), "]") + 1)
                    AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex) & vbCrLf
                    Message(portIndex) = ""
                    HasHeader = True
                Case "R"
                    TMessage(portIndex) = Mid(Message(portIndex), InStr(1, Message(portIndex), "["), InStr(1, Message(portIndex), "]") + 1)
                    AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex) & vbCrLf
                    Message(portIndex) = ""
                Case "RS"
                    TMessage(portIndex) = Mid(Message(portIndex), InStr(1, Message(portIndex), "["), InStr(1, Message(portIndex), "]") + 1)
                    AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex) & vbCrLf
                    Message(portIndex) = ""
                Case "TU"
                    TMessage(portIndex) = Mid(Message(portIndex), InStr(1, Message(portIndex), "["), InStr(1, Message(portIndex), "]") + 1)
                    AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex) & vbCrLf
                    Message(portIndex) = ""
                Case "T"
                    TMessage(portIndex) = Mid(Message(portIndex), InStr(1, Message(portIndex), "["), InStr(1, Message(portIndex), "]") + 1)
                    AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
                    If HasHeader Then processData AllMessage(portIndex)
                    AllMessage(portIndex) = ""
                    TMessage(portIndex) = ""
                    Message(portIndex) = ""
                    HasHeader = False
                    lblStatus.Caption = "Ready to receive data..."
                Case Else
                    TMessage(portIndex) = Mid(Message(portIndex), InStr(1, Message(portIndex), "["), InStr(1, Message(portIndex), "]") + 1)
                    AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex) & vbCrLf
                    Message(portIndex) = ""
            End Select
        End If
    End If
End If
If MessageFormat = "ILP" Then
    lblStatus.Caption = "Receiving Data..."
    If WaitCheckSum Then
        If Len(Message(portIndex)) = 2 Then
            TMessage(portIndex) = TMessage(portIndex) & Message(portIndex) & vbCrLf
            AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
            dateStart(portIndex) = Now
            Do While DateDiff("s", dateStart(portIndex), Now) <= 5
                'wait 3 sec.
            Loop
            mscMain(portIndex).Output = AckStr
            TMessage(portIndex) = ""
            Message(portIndex) = ""
            WaitCheckSum = False
            If ProcessNow Then
                processData AllMessage(portIndex)
                ProcessNow = False
            End If
        End If
    End If
    If InStr(1, Message(portIndex), vbCrLf) > 0 Then
        If Mid(Message(portIndex), 1, 1) = ":" Then
            WaitCheckSum = False
            TMessage(portIndex) = ""
            TMessage(portIndex) = TMessage(portIndex) & Message(portIndex)
            AckStr = "!" & Mid(Message(portIndex), 2, 4)
            Message(portIndex) = ""
            ProcessNow = False
        Else
            If LastMessage = "" Then
                LastMessage = String(4 - Len(Trim(Str(Val("&H" & Mid(Message(portIndex), 1, 2)) - 1))), "0") & Trim(Str(Val("&H" & Mid(Message(portIndex), 1, 2)) - 1))
            End If
            TMessage(portIndex) = TMessage(portIndex) & Message(portIndex)
            If Left(Right(Message(portIndex), 4), 2) = "21" Then ProcessNow = True
            Message(portIndex) = ""
            WaitCheckSum = True
        End If
    End If
End If
If MessageFormat = "ILP2" Then
    lblStatus.Caption = "Receiving Data..."
    If InStr(1, Message(portIndex), vbCrLf) > 0 Then
        If getAcTmsgType(Message(portIndex)) = "H" Then
            If Len(Trim(AllMessage(portIndex))) = 0 Then
                TMessage(portIndex) = ""
                AllMessage(portIndex) = Message(portIndex)
                Message(portIndex) = ""
                sendAck portIndex
            Else
                TMessage(portIndex) = Message(portIndex)
                AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
                processData AllMessage(portIndex)
                lblStatus.Caption = "Ready to receive data..."
                AllMessage(portIndex) = ""
                TMessage(portIndex) = ""
                Message(portIndex) = ""
                sendAck portIndex
            End If
        ElseIf getAcTmsgType(Message(portIndex)) = "L" Then
            TMessage(portIndex) = Message(portIndex)
            AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
            processData AllMessage(portIndex)
            lblStatus.Caption = "Ready to receive data..."
            AllMessage(portIndex) = ""
            TMessage(portIndex) = ""
            Message(portIndex) = ""
            sendAck portIndex
        Else
            TMessage(portIndex) = Message(portIndex)
            AllMessage(portIndex) = AllMessage(portIndex) & TMessage(portIndex)
            Message(portIndex) = ""
            sendAck portIndex
        End If
    End If

End If
End Sub

Private Sub processData(tMess As String, Optional mIndex As Integer = 0)
Dim ctr As Integer
Dim myDate As String
Dim arrTemp() As String
If isDBconnected Then
    If InStr(1, tMess, "LOT#") > 0 And InStr(1, tMess, "UCNAME") > 0 Then
    Else
        If MsgDef.InstrumentID = "RAPID" Then
            eConnection.Execute "Insert Into LabUnprocessed (ResultData, InstrumentID) select '" & Replace(tMess, "'", "''") & "','" & InstrumentID & "'"
            displayresults MsgDef.textDisplay(tMess)
            If Autoprocess Then
                frmProcess.Show
            End If
            If EnableDTP = 1 And Val(GetSetting("Lab-I", "Settings", "Autoprocess", "0")) = 1 Then
                ObjDTP.ProcessDTP MsgDef.segment_OBX(tMess), MsgDef.segment_oBI(tMess), tMess
                If ObjDTP.ErrorNum <> 0 Then
                    displayresults ObjDTP.ErrorDesc
                End If
            End If
            If Val(GetSetting("Lab-I", "HL7", "Enable", "0")) = 1 Then
                If Val(GetSetting("LAB-i", "HL7 Settings", "AutoWriteResult", 0)) = 1 Then
                    WriteToFile HL7_ORU(tMess, rstExam2), GetSetting("LAB-i", "HL7 Settings", "FWDIR", App.Path & "\HostSpooler")
                End If
            End If
        ElseIf MessageFormat = "STKS" Then
            ResultOtherInfo = get_OtherResultInfo(tMess)
            If Val(GetSetting("Lab-I", "Settings", "Autoprocess", "0")) = 1 Then
                myDate = SaveResult(segment_OBX2(tMess, rstExam2))
                If SaveResultFailed Then
                    'Save failed and was rolled back: keep the raw message so it can be reprocessed
                    eConnection.Execute "Insert Into LabUnprocessed (ResultData, InstrumentID) select '" & Replace(tMess, "'", "''") & "','" & InstrumentID & "'"
                    displayresults "Result not saved - message kept in LabUnprocessed for reprocessing."
                ElseIf Val(GetSetting("Lab-I", "Settings", "Autoprint", "0")) = 1 Then
                    If Len(ResultOtherInfo(2)) Then eConnection.Execute "Insert Into LabTempPrint (InstrumentID, SpecimenID, PatientName, Resultdate) select '" & InstrumentID & "','" & ResultOtherInfo(6) & "','" & ResultOtherInfo(2) & "','" & myDate & "'"
                End If
            Else
                eConnection.Execute "Insert Into LabUnprocessed (ResultData, InstrumentID) select '" & Replace(tMess, "'", "''") & "','" & InstrumentID & "'"
            End If
            displayresults DisplayResult2Text(segment_OBX(tMess, rstExam2), ResultOtherInfo)
            If EnableDTP = 1 And Val(GetSetting("Lab-I", "Settings", "Autoprocess", "0")) = 1 Then
                ObjDTP.ProcessDTP segment_OBX2(tMess, rstExam2), get_OtherResultInfo(tMess, 1), tMess
                If ObjDTP.ErrorNum <> 0 Then
                    displayresults ObjDTP.ErrorDesc
                End If
            End If
            If Val(GetSetting("Lab-I", "HL7", "Enable", "0")) = 1 Then
                If Val(GetSetting("LAB-i", "HL7 Settings", "AutoWriteResult", 0)) = 1 Then
                    WriteToFile HL7_ORU(tMess, rstExam2), GetSetting("LAB-i", "HL7 Settings", "FWDIR", App.Path & "\HostSpooler")
                End If
            End If
       
        
        Else
            eConnection.Execute "Insert Into LabUnprocessed (ResultData, InstrumentID) select '" & Replace(tMess, "'", "''") & "','" & InstrumentID & "'"
            displayresults MsgDef.textDisplay(tMess)
            If Autoprocess Then
                frmProcess.Show
            End If
            If EnableDTP = 1 And Val(GetSetting("Lab-I", "Settings", "Autoprocess", "0")) = 1 Then
                ObjDTP.ProcessDTP MsgDef.segment_OBX(tMess), MsgDef.segment_oBI(tMess), tMess
                If ObjDTP.ErrorNum <> 0 Then
                    displayresults ObjDTP.ErrorDesc
                End If
            End If
            If Val(GetSetting("Lab-I", "HL7", "Enable", "0")) = 1 Then
                If Val(GetSetting("LAB-i", "HL7 Settings", "AutoWriteResult", 0)) = 1 Then
                    WriteToFile HL7_ORU(tMess, rstExam2), GetSetting("LAB-i", "HL7 Settings", "FWDIR", App.Path & "\HostSpooler")
                End If
            End If
        End If
    End If
Else
    displayresults "No Connection to Database"
End If
End Sub

Private Sub rtbRcvDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 192 Then
    showConsole
End If
End Sub

Private Sub consoleCmd(myCommand As String)
Dim mCommand As String
Dim ctr As Integer
mCommand = Trim(UCase(myCommand))
Select Case mCommand
    Case "OPENDB"
        If isDBconnected Then
            displayresults "PROGRAM IS ALREADY CONNECTED."
            displayresults "DATABASE SERVER:" & DBServerType
        Else
            OpenConnection
            mdiMain.mnuConnServ.Visible = Not isDBconnected
            displayresults "CONNECTED TO SERVER."
        End If
    Case "CLS"
        displayresults "", True
    Case "DISCONNECT"
    Case "CLOSEDB"
        If isDBconnected Then
            Unload frmSearchResult
            isResultVisible = False
            CloseConnection
            mdiMain.mnuConnServ.Visible = Not isDBconnected
            displayresults "SERVER DISCONNECTED."
        Else
            displayresults "CONNECTION ALREADY CLOSED."
        End If
    Case "VER"
        displayresults App.Title
        displayresults "Standard Edition"
        displayresults "VERSION: " & App.Major & "." & App.Minor & "." & App.Revision
        displayresults "Marsman Drysdale Medical Products, Inc."
    Case "CLEARBUFFER"
    Case "RESETRX"
    Case "HELP"
            displayresults "~*~COMMANDS~*~"
            displayresults "CLS"
            displayresults "DISCONNECT"
            displayresults "VER"
            displayresults "HELP"
            displayresults "CONFIG"
            displayresults "SETTINGS"
            displayresults "LOOPBACK"
            displayresults ""
    Case "CONFIG"
        For ctr = 1 To 4
            displayresults "Control # " & ctr
            displayresults "SERIAL PORT NUMBER: " & mscMain(ctr - 1).CommPort
            displayresults "SETTINGS: " & mscMain(ctr - 1).Settings
            displayresults "HANDSHAKING: " & mscMain(ctr - 1).Handshaking
            displayresults "INBUFFERSIZE: " & mscMain(ctr - 1).InBufferSize
            displayresults "OUTBUFFERSIZE: " & mscMain(ctr - 1).OutBufferSize
            displayresults "RTHRESHOLD: " & mscMain(ctr - 1).RThreshold
            displayresults "STHRESHOLD: " & mscMain(ctr - 1).SThreshold
            displayresults "INPUTLEN: " & mscMain(ctr - 1).InputLen
            displayresults "PARITY REPLACE: " & mscMain(ctr - 1).ParityReplace
            displayresults "EOF ENABLE: " & mscMain(ctr - 1).EOFEnable
            displayresults "NULL DISCARD: " & mscMain(ctr - 1).NullDiscard
            displayresults "RTS ENABLE: " & mscMain(ctr - 1).RTSEnable
            displayresults "DTR ENABLE: " & mscMain(ctr - 1).DTREnable
            displayresults ""
        Next ctr
    Case "SETTINGS"
        displayresults "Diff Count Total: " & GetSetting("Lab-I", "Settings", "DiffCount", "1")
        displayresults "Database: " & DBDatabase
    Case "LOOPBACK"
        If LoopBack Then
            LoopBack = False
            displayresults "Loopback  Mode OFF"
        Else
            LoopBack = True
            displayresults "Loopback  Mode ON"
        End If
    Case Else
        If Mid(mCommand, 1, 7) = "SAVELOG" Then
             SaveLog getStringData(mCommand, "SAVELOG", Format(Now, "mmddyy"))
        ElseIf Mid(mCommand, 1, 7) = "LOADLOG" Then
             LoadLog getStringData(mCommand, "LOADLOG", Format(Now, "mmddyy")) & ".LOG"
        ElseIf Mid(mCommand, 1, 7) = "SETDIFF" Then
            SaveSetting "Lab-I", "Settings", "DiffCount", getStringData(mCommand, "SETDIFF", "1")
            displayresults "Diff Count Total: " & GetSetting("Lab-I", "Settings", "DiffCount", "1")
        ElseIf Mid(mCommand, 1, 9) = "ADDREPORT" Then
            If Len(Trim(getStringData(mCommand, "ADDREPORT", ""))) > 0 Then
                eConnection.Execute "Insert Into tbLabReportNames (rptName, rptDefault) Select '" & Trim(getStringData(mCommand, "ADDREPORT", "")) & "',False"
                displayresults "Added " & Trim(getStringData(mCommand, "ADDREPORT", ""))
            End If
        ElseIf Mid(mCommand, 1, 4) = "SEND" Then
            mdiMain.wsBroadcast.SendData getStringData(mCommand, "SEND", "")
        Else
            If Len(Trim(mCommand)) > 0 Then displayresults "UNKNOWN COMMAND"
        End If
End Select
End Sub

'Monthly buffer cleanup helper. LabUnprocessed has no date column, and rows still in it
'are results that were never saved, so they are written to
'<data folder>\UnprocessedArchive\<InstrumentID>_<timestamp>.txt first. Only the rows that
'were written are deleted; if anything fails, nothing is deleted.
Private Sub ArchiveUnprocessed()
Dim rstTemp As ADODB.Recordset
Dim tsOut As Object
Dim mFolder As String
Dim maxID As Long
On Error GoTo errFound
If DBServerType = "LOCAL DATABASE" Then
    mFolder = Replace(DBDatabase, "Laboratory.mdb", "") & "UnprocessedArchive"
Else
    mFolder = App.Path & "\UnprocessedArchive"
End If
maxID = -1
Set rstTemp = myConnection.Execute("select ResultID, ResultData from LabUnprocessed where InstrumentID = '" & InstrumentID & "' order by ResultID")
If Not rstTemp.EOF Then
    If Not fsoFiles.FolderExists(mFolder) Then fsoFiles.CreateFolder mFolder
    Set tsOut = fsoFiles.OpenTextFile(mFolder & "\" & InstrumentID & "_" & Format(Now, "yyyymmdd_hhnnss") & ".txt", 8, True)
    Do While Not rstTemp.EOF
        tsOut.WriteLine "===== ResultID " & rstTemp.Fields("ResultID").Value & " ====="
        tsOut.WriteLine "" & rstTemp.Fields("ResultData").Value
        maxID = CLng(rstTemp.Fields("ResultID").Value)
        rstTemp.MoveNext
    Loop
    tsOut.Close
    Set tsOut = Nothing
End If
rstTemp.Close
Set rstTemp = Nothing
If maxID >= 0 Then
    myConnection.Execute "delete from LabUnprocessed where InstrumentID = '" & InstrumentID & "' and ResultID <= " & maxID
    displayresults "Unprocessed results saved to " & mFolder
End If
Exit Sub
errFound:
    displayresults "Auto Archive: unprocessed results could not be saved (" & Err.Description & "). Buffer not cleared."
End Sub
Private Sub tmHL7_Timer()
Dim mFilename As String
Dim mDirectory As String
Dim mcount As Long
mDirectory = GetSetting("LAB-i", "HL7 Settings", "FRDIR", App.Path & "\Received")
If Val(GetSetting("ComSet", "Control 1", "FullHandshake", "0")) = 1 Then
    '-------Delete Buffer
    If Trim(GetSetting("LAB-i", "AutoArchive", "LastArchive", "NONE")) <> Trim(Format(Now, "yyyymm")) Then
        displayresults "Auto Archive Started..."
        ArchiveUnprocessed
        myConnection.Execute "delete from tblabRequest where val(left(messagedate,6))  < " & Format(DateAdd("m", -3, Now), "yyyymm")
        displayresults "Closing Database..."
        CloseConnection
        displayresults "Compressing Database..."
        DAO.DBEngine.CompactDatabase Replace(DBDatabase, "Laboratory.mdb", "") & "LabBuffer.mdb", Replace(DBDatabase, "Laboratory.mdb", "") & "tLabBuffer.mdb", dbLangGeneral, , ";pwd=" & DBPassword
        DAO.DBEngine.CompactDatabase Replace(DBDatabase, "Laboratory.mdb", "") & "LabRequestBuffer.mdb", Replace(DBDatabase, "Laboratory.mdb", "") & "tLabRequestBuffer.mdb", dbLangGeneral, , ";pwd=" & DBPassword
        fsoFiles.CopyFile Replace(DBDatabase, "Laboratory.mdb", "") & "tLabBuffer.mdb", Replace(DBDatabase, "Laboratory.mdb", "") & "LabBuffer.mdb", True
        fsoFiles.CopyFile Replace(DBDatabase, "Laboratory.mdb", "") & "tLabRequestBuffer.mdb", Replace(DBDatabase, "Laboratory.mdb", "") & "LabRequestBuffer.mdb", True
        fsoFiles.DeleteFile Replace(DBDatabase, "Laboratory.mdb", "") & "tLabBuffer.mdb", True
        fsoFiles.DeleteFile Replace(DBDatabase, "Laboratory.mdb", "") & "tLabRequestBuffer.mdb", True
        displayresults "Initializing Database..."
        OpenConnection
        SaveSetting "LAB-i", "AutoArchive", "LastArchive", Trim(Format(Now, "yyyymm"))
        displayresults "Archive Complete."
    End If

    If Len(Trim(GetSetting("LAB-i", "HL7 Settings", "FREXT", "txt"))) = 0 Then
        mFilename = Dir(GetSetting("LAB-i", "HL7 Settings", "FRDIR", App.Path & "\Received") & "\*")
    Else
        mFilename = Dir(GetSetting("LAB-i", "HL7 Settings", "FRDIR", App.Path & "\Received") & "\*." & GetSetting("LAB-i", "HL7 Settings", "FREXT", "txt"))
    End If
    mcount = fsoFiles.GetFolder(mDirectory & "\").Files.Count
    If mcount > 10 Then tmHL7.Interval = 100 Else tmHL7.Interval = 6000
    Debug.Print mcount & ": " & tmHL7.Interval
    If mcount > 0 Then
        displayresults "Processing Request Message:"
        displayresults SaveHL7toDB(GetSetting("LAB-i", "HL7 Settings", "FRDIR", App.Path & "\Received"), GetSetting("LAB-i", "HL7 Settings", "FREXT", "txt"))
    End If
End If
End Sub

Private Sub tmrAutoPrint_Timer()
On Error GoTo errFound
Dim rstTemp As ADODB.Recordset
If Val(GetSetting("Lab-I", "Settings", "Autoprocess", "0")) = 1 Then
    Set rstTemp = eConnection.Execute("Select * from LabTempPrint where InstrumentID = '" & InstrumentID & "'")
    If rstTemp.RecordCount > 0 Then
        rstTemp.MoveFirst
        Do While Not rstTemp.EOF
            If Not frmProcess.Visible Then
                HostSA.ResultAutoPrint GetSetting("Lab-I", "Settings", "ProfileID", "1"), InstrumentID, rstTemp.Fields("SpecimenID").Value, rstTemp.Fields("PatientName").Value, rstTemp.Fields("ResultDate").Value
                eConnection.Execute "Delete from LabTempPrint where PrintID = " & rstTemp.Fields("PrintID").Value
                rstTemp.MoveNext
            End If
        Loop
    End If
End If
errFound:
    If Err.Number = -2147217865 Then
        FixLink DBDatabase
    ElseIf Err.Number = 0 Then
    Else
        MsgBox Err.Number & ": " & Err.Description, vbOKOnly, "Error"
    End If
End Sub

Private Sub tmrStatus_Timer()
DoEvents
    checkComStatus
    If Val(GetSetting("Lab-I", "Settings", "Autoprint", "0")) = 1 Then tmrAutoPrint.Interval = 1000 Else tmrAutoPrint.Interval = 1000
End Sub

Private Sub tmrWorklist_Timer()
Dim rstTemp As ADODB.Recordset
DoEvents
If Val(GetSetting("ComSet", "Control 1", "FullHandshake", "0")) = 1 And Val(GetSetting("ComSet", "Control 1", "Enabled", "0")) = 1 Then
    Set rstTemp = myConnection.Execute("Select top 1 * from tblabRequest where messageStatus = '1'")
    If rstTemp.RecordCount = 1 Then
        If MessageFormat = "STKS" Then
            If Not ReceiveMode Then
                If Not WaitEnq Then
                    sendEnq 0
                    WaitEnq = True
                    WLTimer = 0
                    SendMode = True
                    dFilename = ""
                Else
                    If Not WaitAck Then
                        WLTimer = WLTimer + 1
                        If WLTimer >= 4 Then WaitEnq = False
                        If Syned = True Then WaitEnq = False
                        SendMode = False
                    End If
                End If
            End If
        ElseIf MessageFormat = "ASTM" Then
            If Not ReceiveMode Then
                If Not WaitAck Then
                    sendEnq 0
                    WaitEnq = False
                    WaitAck = True
                    WLTimer = 0
                    SendMode = True
                    dFilename = ""
                    ProcessRequest = False
                Else
                    If Syned Then
                        WLTimer = WLTimer + 1
                        If WLTimer >= 4 Then WaitAck = False
                        If Syned = True Then Syned = False
                        SendMode = False
                    ElseIf WLTimer >= 4 Then WaitAck = False
                    End If
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub txtCommand_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    consoleCmd txtCommand.Text
    txtCommand.Text = ""
End If
If KeyCode = 192 Then
    showConsole
End If
If KeyCode = vbKeyF2 Then
    frmCharmap.Show 1
    txtCommand.SelText = ReturnValue
End If
If Shift = 1 And KeyCode = vbKeyF12 Then
    If txtCommand.PasswordChar = "*" Then txtCommand.PasswordChar = "" Else txtCommand.PasswordChar = "*"
End If
End Sub

Private Sub displayresults(Optional displayText As String = "", Optional Clear As Boolean = False, Optional NoLF As Boolean = False)
On Error GoTo errFound
If Clear Then
    rtbRcvDisplay.Text = ""
Else
    If Len(Trim(displayText)) > 0 Then
        rtbRcvDisplay.Text = rtbRcvDisplay.Text & displayText & IIf(NoLF, "", vbCrLf)
    Else
        rtbRcvDisplay.Text = rtbRcvDisplay.Text & displayText & IIf(NoLF, "", vbCrLf)
    End If
End If
If Len(rtbRcvDisplay.Text) > 20000 Then
    rtbRcvDisplay.Text = Mid(rtbRcvDisplay.Text, 20000)
End If
rtbRcvDisplay.SelStart = Len(rtbRcvDisplay.Text)
errFound:
    If Err.Number <> 0 Then
        displayresults "", True
        Resume Next
    End If
End Sub
Private Sub SaveLog(filename As String)
Dim notFound As Boolean
Dim ctr As Integer
Dim NewFileName
On Error GoTo errFound
ctr = 1
notFound = False
Do While notFound = False
    If Len(Trim(Dir(App.Path & "\" & filename & "(" & ctr & ").log"))) > 0 Then
        notFound = False
    Else
        notFound = True
        NewFileName = App.Path & "\" & filename & "(" & ctr & ").log"
    End If
    ctr = ctr + 1
Loop
rtbRcvDisplay.SaveFile NewFileName, rtfText
errFound:
If Err.Number > 0 Then
    displayresults "ERROR: " & Err.Description
End If
    displayresults "LOGFILENAME: " & NewFileName
End Sub

Private Sub LoadLog(filename As String)
    If Len(Trim(Dir(App.Path & "\" & filename))) > 0 Then
        rtbRcvDisplay.LoadFile App.Path & "\" & filename, rtfText
    Else
        displayresults "LOG FILE NOT FOUND"
    End If
End Sub

Private Function PartMessage(Message As String) As String
Dim ctr As Double
Dim tempstr As String
Dim start As Boolean
Dim strEnd As Boolean
start = False
strEnd = False
tempstr = ""
For ctr = 1 To Len(Message)
    If Mid(Message, ctr, 1) = Chr(2) Then start = True
    If start Then
        tempstr = tempstr + Mid(Message, ctr, 1)
        If Mid(Message, ctr, 1) = Chr(3) Then
            strEnd = True
            Exit For
        End If
    End If
Next ctr
If strEnd Then
'invalid procedure call
    tempstr = Mid(tempstr, 4, Len(tempstr) - 8)
    PartMessage = tempstr
Else
    PartMessage = ""
End If
End Function

Private Sub checkComStatus()
Dim ctr As Integer
Dim tempstr As String
Dim tempStr2 As String
Dim hasConnected As Boolean

hasConnected = False
tempstr = "Receive Data - "
For ctr = 1 To 4
If DetectCon = 1 Then
    If (mscMain(ctr - 1).DSRHolding = False And mscMain(ctr - 1).CTSHolding = False) Then
'        tempStr = tempStr & "C" & ctr & " "
        If hasConnected = False Then lblStatus.Caption = "Data Cable not Connected or DMS not Ready..."
        lblTitle.Caption = tempstr '"Receive Data - [Disconnected]"
    ElseIf (mscMain(ctr - 1).DSRHolding = False And mscMain(ctr - 1).CTSHolding = True) Then
'        tempStr = tempStr & "UD" & ctr
        If hasConnected = False Then lblStatus.Caption = "Instrument or DMS not detected."
        lblTitle.Caption = tempstr '"Receive Data - [Disconnected]"
    Else
        hasConnected = True
        tempstr = tempstr & "C" & ctr & " "
        tempStr2 = tempStr2 & "C" & ctr & " "
        lblStatus.Caption = "Ready to receive data from - " & tempStr2
        lblTitle.Caption = tempstr '"Receive Data - [Connected]"
    End If
Else
    tempstr = tempstr & "C" & ctr & " "
    lblStatus.Caption = "Ready to receive data..."
    lblTitle.Caption = "Receive Data - [Connected]"
End If
Next ctr
End Sub

Private Sub OpenLocalConnection()
    Set eConnection = myConnection
End Sub

Private Function getConStr() As String
Dim tempstr As String
Dim mUserName As String
Dim mDatabase As String
Dim mPassword  As String
    tempstr = GetSetting("lab-i", "LOCAL", "UserName", "")
    If Len(Trim(tempstr)) > 0 Then
        mUserName = Blowfish.DecryptString(tempstr, myKey, True)
    Else
        mUserName = ""
    End If
    tempstr = GetSetting("lab-i", "LOCAL", "DATABASEPASSWORD", "")
    If Len(Trim(tempstr)) > 0 Then
        mPassword = Blowfish.DecryptString(tempstr, myKey, True)
    Else
        mPassword = ""
    End If
    tempstr = GetSetting("lab-i", "LOCAL", "DATABASE", "")
    If Len(Trim(tempstr)) > 0 Then
        mDatabase = Blowfish.DecryptString(tempstr, myKey, True)
    Else
        mDatabase = ""
    End If
    getConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mDatabase & ";User Id=" & mUserName & ";Password=;" & "Jet OLEDB:Database Password=" & mPassword & ";"
End Function

Private Sub saveUcontrol(tMess As String)
Dim lotNum As String
Dim ucName As String
Dim expiration As String
Dim RunNum As String
Dim ctr As Integer
Dim tmpResult As Variant
Dim myResults2 As Variant
Dim temps As String
lotNum = getMessageValues(tMess, "LOT#")
ucName = getMessageValues(tMess, "UCNAME")
expiration = getMessageValues(tMess, "EXDATE")
RunNum = getMessageValues(tMess, "RUN#")

tmpResult = Split(segment_OBX(tMess, rstExam2), vbCrLf)
For ctr = 0 To UBound(tmpResult)
    myResults2 = Split(tmpResult, "|")
    temps = "Insert into (InstrumentID, resultDate, LOTNUM, UCNAME, ExamName, Result, Unit, Status, RunNum)" & vbCrLf & _
        "Select '" & InstrumentID & "',#" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy HH:mm") & "#,'" & lotNum & "','" & ucName & "','" & myResults2(3) & "'," & myResults2(5) & ",'" & myResults2(6) & "','0','" & RunNum & "';"
Next ctr
End Sub

Private Sub clearPublicVar()
Dim ctr As Integer
For ctr = 0 To 3
    Message(ctr) = ""
    TMessage(ctr) = ""
    AllMessage(ctr) = ""
    TotalMsgCount(ctr) = 0
    MsgCount(ctr) = 0
Next ctr
End Sub

Private Function getILPmsgType(mStr As String) As String
Dim tempstr As String
tempstr = Format(Mid(mStr, 3, 1), ">")
If InStr(1, "HPORLQ", tempstr) > 0 Then
    getILPmsgType = tempstr
Else
    getILPmsgType = "U"
End If
End Function

Private Function getASTMmsgType(mStr As String) As String
Dim tempstr As String
tempstr = Format(Mid(mStr, 3, 1), ">")
If InStr(1, "HPORLQ", tempstr) > 0 Then
    getASTMmsgType = tempstr
Else
    getASTMmsgType = "U"
End If
End Function

Private Function getAcTmsgType(mStr As String) As String
Dim tempstr As String
tempstr = Format(Mid(mStr, 3, 1), ">")
If InStr(1, "HPORLQ", tempstr) > 0 Then
    getAcTmsgType = tempstr
Else
    getAcTmsgType = "U"
End If
End Function

Private Function getCXSYSmsgType(mStr As String) As String
Dim tempstr() As String
tempstr = Split(mStr, ",")
Select Case tempstr(2)
    Case "01": getCXSYSmsgType = "H"
    Case "03": getCXSYSmsgType = "R"
    Case "05": getCXSYSmsgType = "T"
    Case "07": getCXSYSmsgType = "LCAL"
    Case "09": getCXSYSmsgType = "MCAL"
    Case "11": getCXSYSmsgType = "RS"
    Case "13": getCXSYSmsgType = "TU"
    Case "21": getCXSYSmsgType = "Reagent Pack Header"
    Case "23": getCXSYSmsgType = "Result for calibration"
    Case "25": getCXSYSmsgType = "Expanded result for Calibration"
    Case "27": getCXSYSmsgType = "End of reagent pack"
    Case "81": getCXSYSmsgType = "Expanded result"
    Case "83": getCXSYSmsgType = "Expanded multipoint calibration result"
    Case Else
End Select
End Function

Private Sub sendAck(Optional mIndex As Integer = 0)
    mscMain(mIndex).Output = Chr(6)
End Sub

Private Sub sendEnq(Optional mIndex As Integer = 0)
    mscMain(mIndex).Output = Chr(5)
End Sub

'Private Function SaveResult(msgHL7format As Variant) As String
'Dim ctr1 As Integer
'Dim counter As Integer
'Dim myResults As Variant
'Dim myResults2 As Variant
'Dim strSql As String
'Dim rstTemp As ADODB.Recordset
'Dim SampleNum As String
'Dim rstSearch As New ADODB.Recordset
'Dim temps As String
'Dim tempResultNum As String
'Dim myResultValue As String
'Dim SaveHeader As Boolean
'Dim myTempTime As String
'Dim tPatientName, tSpecimenID, tDate As String
'SaveHeader = False
'tPatientName = ""
'tSpecimenID = ""
'tDate = ""
''On Error GoTo errFound
'myTempTime = Format(Now, "HH:mm:ss")
'myResults = Split(msgHL7format, vbCr)
'counter = UBound(myResults) - 1
'If rstExam2.RecordCount > 0 Then
'rstExam2.MoveFirst
'myConnection.BeginTrans
'For ctr1 = 0 To counter
'    myResults2 = Split(myResults(ctr1), "|")
'    If Len(ResultOtherInfo(3)) > 3 Then SampleNum = ResultOtherInfo(3) Else SampleNum = String(3 - Len(ResultOtherInfo(3)), "0") & ResultOtherInfo(3)
'        If ProcessDateOnly = 1 Then
'            If DBServerType = "MYSQL SERVER" Then
'                temps = "Select * from tbLabInstrumentResults where InstrumentID = '" & _
'                    InstrumentID & "' and PatientName = '" & ResultOtherInfo(2) & "' and AccessionNumOrg = '" & _
'                    ResultOtherInfo(6) & "' and format(Resultdate,'mm/dd/yyyy') = '" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy") & "' and ExamName = '" & myResults2(3) & "'"
'                SaveResult = formatDT(myResults2(14), 3)
'            Else
'                temps = "Select * from tbLabInstrumentResults where InstrumentID = '" & _
'                    InstrumentID & "' and PatientName = '" & ResultOtherInfo(2) & "' and AccessionNumOrg = '" & _
'                    ResultOtherInfo(6) & "' and format(Resultdate,'mm/dd/yyyy') = '" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy") & "' and ExamName = '" & myResults2(3) & "'"
'                SaveResult = formatDT(myResults2(14), 3)
'            End If
'        Else
'            If DBServerType = "MYSQL SERVER" Then
'                temps = "Select * from tbLabInstrumentResults where InstrumentID = '" & _
'                    InstrumentID & "' and PatientName = '" & ResultOtherInfo(2) & "' and AccessionNumOrg = '" & _
'                    ResultOtherInfo(6) & "' and Resultdate = '" & Format(formatDT(myResults2(14), 3), "yyyy-mm-dd HH:mm:ss") & "' and ExamName = '" & myResults2(3) & "'"
'                SaveResult = formatDT(myResults2(14), 3)
'            Else
'
'                If Val(GetSetting("Lab-I", "Settings", "UseCurrentTime", "0")) = 1 Then
'                    temps = "Select a.*, b.* from tbLabInstrumentResults a, tbLabInstrumentRValues b where a.ResultID = b.ResultID and a.InstrumentID = '" & _
'                        InstrumentID & "' and a.PatientName = '" & ResultOtherInfo(2) & "' and a.AccessionNumOrg = '" & _
'                        ResultOtherInfo(6) & "' and format(a.Resultdate,'mm/dd/yyyy HH:mm') = '" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy ") & myTempTime & "' and b.ExamName = '" & myResults2(3) & "'"
'
'                        SaveResult = formatDT(myResults2(14), 3)
'                Else
'                    temps = "Select a.*, b.* from tbLabInstrumentResults a, tbLabInstrumentRValues b where a.ResultID = b.ResultID and a.InstrumentID = '" & _
'                        InstrumentID & "' and a.PatientName = '" & ResultOtherInfo(2) & "' and a.AccessionNumOrg = '" & _
'                        ResultOtherInfo(6) & "' and format(a.Resultdate,'mm/dd/yyyy HH:mm') = '" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy HH:mm") & "' and b.ExamName = '" & myResults2(3) & "'"
'
'                    SaveResult = formatDT(myResults2(14), 3)
'                End If
'            End If
'        End If
'    Set rstSearch = myConnection.Execute(temps)
'    'If saveStringResult Then myResultValue = "'" & IIf(Len(Trim(myResults2(5))) = 0, "0", myResults2(5)) & "'" Else myResultValue = IIf(Len(Trim(myResults2(5))) = 0, "0", myResults2(5))
'    myResultValue = "'" & IIf(Len(Trim(myResults2(5))) = 0, "0", myResults2(5)) & "'"
'    If rstSearch.RecordCount > 0 Then
'        If ProcessDateOnly = 1 Then
'            If DBServerType = "MYSQL SERVER" Then
'                temps = "Update tbLabInstrumentResults SET Result = " & myResultValue & " " & vbCrLf & _
'                    "where Status = '0' and InstrumentID = '" & _
'                    InstrumentID & "' and PatientName = '" & ResultOtherInfo(2) & "' and AccessionNumOrg = '" & _
'                    ResultOtherInfo(6) & "' and format(Resultdate,'mm/dd/yyyy') = #" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy") & "# and ExamName = '" & myResults2(3) & "'"
'                SaveResult = formatDT(myResults2(14), 3)
'            Else
'                temps = "Update tbLabInstrumentResults SET Result = " & myResultValue & " " & vbCrLf & _
'                    "where Status = '0' and InstrumentID = '" & _
'                    InstrumentID & "' and PatientName = '" & ResultOtherInfo(2) & "' and AccessionNumOrg = '" & _
'                    ResultOtherInfo(6) & "' and format(Resultdate,'mm/dd/yyyy') = #" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy") & "# and ExamName = '" & myResults2(3) & "'"
'                SaveResult = formatDT(myResults2(14), 3)
'            End If
'        Else
'            If DBServerType = "MYSQL SERVER" Then
'                temps = "Update tbLabInstrumentResults SET Result = " & myResultValue & " " & vbCrLf & _
'                    "where Status = '0' and InstrumentID = '" & _
'                    InstrumentID & "' and PatientName = '" & ResultOtherInfo(2) & "' and AccessionNumOrg = '" & _
'                    ResultOtherInfo(6) & "' and Resultdate = '" & Format(formatDT(myResults2(14), 3), "yyyy-mm-dd HH:mm:ss") & "' and ExamName = '" & myResults2(3) & "'"
'                SaveResult = formatDT(myResults2(14), 3)
'            Else
'                temps = "Update tbLabInstrumentResults SET Result = " & myResultValue & " " & vbCrLf & _
'                    "where Status = '0' and InstrumentID = '" & _
'                    InstrumentID & "' and PatientName = '" & ResultOtherInfo(2) & "' and AccessionNumOrg = '" & _
'                    ResultOtherInfo(6) & "' and Resultdate = #" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy HH:mm:ss") & "# and ExamName = '" & myResults2(3) & "'"
'                SaveResult = formatDT(myResults2(14), 3)
'            End If
'        End If
'    Else
'        If ProcessDateOnly = 1 Then
'            If DBServerType = "MYSQL SERVER" Then
'                temps = "Insert into tbLabInstrumentResults (InstrumentID, Resultdate, Accessionnum, PatientName, ExamName, Result, Unit, Status, ProfileExam, AccessionnumOrg, IdNum, BirthDate, SortNum, Room, Age, Sex, Specimen, Doctor)" & vbCrLf & _
'                "Select '" & InstrumentID & "',#" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy") & "#,'" & ResultOtherInfo(6) & "','" & ResultOtherInfo(2) & "','" & myResults2(3) & "'," & myResultValue & ",'" & myResults2(6) & "','0','" & DefaultProfileID & "','" & ResultOtherInfo(6) & "','" & ResultOtherInfo(7) & "','" & _
'                ResultOtherInfo(8) & "'," & myResults2(1) & ",'" & ResultOtherInfo(9) & "','" & ResultOtherInfo(10) & "','" & ResultOtherInfo(11) & "','" & ResultOtherInfo(12) & "','" & ResultOtherInfo(14) & "';"
'
'                SaveResult = formatDT(myResults2(14), 3)
'            Else
'                temps = "Insert into tbLabInstrumentResults (InstrumentID, Resultdate, Accessionnum, PatientName, ExamName, Result, Unit, Status, ProfileExam, AccessionnumOrg, IdNum, BirthDate, SortNum, Room, Age, Sex, Specimen, Doctor)" & vbCrLf & _
'                "Select '" & InstrumentID & "',#" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy") & "#,'" & ResultOtherInfo(6) & "','" & ResultOtherInfo(2) & "','" & myResults2(3) & "'," & myResultValue & ",'" & myResults2(6) & "','0','" & DefaultProfileID & "','" & ResultOtherInfo(6) & "','" & ResultOtherInfo(7) & "','" & _
'                ResultOtherInfo(8) & "'," & myResults2(1) & ",'" & ResultOtherInfo(9) & "','" & ResultOtherInfo(10) & "','" & ResultOtherInfo(11) & "','" & ResultOtherInfo(12) & "','" & ResultOtherInfo(14) & "';"
'
'                SaveResult = formatDT(myResults2(14), 3)
'            End If
'        Else
'            If DBServerType = "MYSQL SERVER" Then
'                temps = "Insert into tbLabInstrumentResults (InstrumentID, Resultdate, Accessionnum, PatientName, ExamName, Result, Unit, Status, ProfileExam, AccessionnumOrg, IdNum, BirthDate, SortNum, Room, Age, Sex, Specimen, Doctor)" & vbCrLf & _
'                "Select '" & InstrumentID & "','" & Format(formatDT(myResults2(14), 3), "yyyy-mm-dd HH:mm:ss") & "','" & ResultOtherInfo(6) & "','" & ResultOtherInfo(2) & "','" & myResults2(3) & "'," & myResultValue & ",'" & myResults2(6) & "','0','" & DefaultProfileID & "','" & ResultOtherInfo(6) & "','" & ResultOtherInfo(7) & "','" & _
'                ResultOtherInfo(8) & "'," & myResults2(1) & ",'" & ResultOtherInfo(9) & "','" & ResultOtherInfo(10) & "','" & ResultOtherInfo(11) & "','" & ResultOtherInfo(12) & "','" & ResultOtherInfo(14) & "';"
'
'                SaveResult = formatDT(myResults2(14), 3)
'            Else
'                If Not SaveHeader Then
'                    If Val(GetSetting("Lab-I", "Settings", "UseCurrentTime", "0")) = 1 Then
'                        temps = "Insert into tbLabInstrumentResults (InstrumentID, Resultdate, Accessionnum, PatientName, Status, ProfileExam, AccessionnumOrg, IdNum, BirthDate, Room, Age, Sex, Specimen, Doctor, ResultID,Operator,Comments)" & vbCrLf & _
'                        "Select '" & InstrumentID & "',#" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy ") & myTempTime & "#,'" & ResultOtherInfo(6) & "','" & ResultOtherInfo(2) & "','0','" & DefaultProfileID & "','" & ResultOtherInfo(6) & "','" & ResultOtherInfo(7) & "','" & _
'                        ResultOtherInfo(8) & "','" & ResultOtherInfo(9) & "','" & ResultOtherInfo(10) & "','" & ResultOtherInfo(11) & "','" & ResultOtherInfo(12) & "','" & ResultOtherInfo(14) & "','" & GenerateResultID(InstrumentID & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy ") & myTempTime & ResultOtherInfo(6) & ResultOtherInfo(2)) & "','" & ResultOtherInfo(17) & "','" & ResultOtherInfo(18) & "';"
'                        tPatientName = ResultOtherInfo(2)
'                        tSpecimenID = ResultOtherInfo(6)
'                        tDate = Format(formatDT(myResults2(14), 3), "mm/dd/yyyy ") & myTempTime
'                    Else
'                        temps = "Insert into tbLabInstrumentResults (InstrumentID, Resultdate, Accessionnum, PatientName, Status, ProfileExam, AccessionnumOrg, IdNum, BirthDate, Room, Age, Sex, Specimen, Doctor, ResultID,Operator, Comments)" & vbCrLf & _
'                        "Select '" & InstrumentID & "',#" & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy HH:mm:ss") & "#,'" & ResultOtherInfo(6) & "','" & ResultOtherInfo(2) & "','0','" & DefaultProfileID & "','" & ResultOtherInfo(6) & "','" & ResultOtherInfo(7) & "','" & _
'                        ResultOtherInfo(8) & "','" & ResultOtherInfo(9) & "','" & ResultOtherInfo(10) & "','" & ResultOtherInfo(11) & "','" & ResultOtherInfo(12) & "','" & ResultOtherInfo(14) & "','" & GenerateResultID(InstrumentID & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy HH:mm:ss") & ResultOtherInfo(6) & ResultOtherInfo(2)) & "','" & ResultOtherInfo(17) & "','" & ResultOtherInfo(18) & "';"
'                        tPatientName = ResultOtherInfo(2)
'                        tSpecimenID = ResultOtherInfo(6)
'                        tDate = Format(formatDT(myResults2(14), 3), "mm/dd/yyyy HH:mm:ss")
'                    End If
'                    myConnection.Execute temps
'                    SaveHeader = True
'                End If
'
'                If Val(GetSetting("Lab-I", "Settings", "UseCurrentTime", "0")) = 1 Then
'                    temps = "Insert into tbLabInstrumentRValues (ExamName, Result, Unit, SortNum, ResultID)" & vbCrLf & _
'                    "Select '" & myResults2(3) & "'," & myResultValue & ",'" & myResults2(6) & "'," & myResults2(1) & ",'" & GenerateResultID(InstrumentID & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy ") & myTempTime & ResultOtherInfo(6) & ResultOtherInfo(2)) & "';"
'
'                    SaveResult = formatDT(myResults2(14), 3)
'                Else
'                    temps = "Insert into tbLabInstrumentRValues (ExamName, Result, Unit, SortNum, ResultID)" & vbCrLf & _
'                    "Select '" & myResults2(3) & "'," & myResultValue & ",'" & myResults2(6) & "'," & myResults2(1) & ",'" & GenerateResultID(InstrumentID & Format(formatDT(myResults2(14), 3), "mm/dd/yyyy HH:mm:ss") & ResultOtherInfo(6) & ResultOtherInfo(2)) & "';"
'
'                    SaveResult = formatDT(myResults2(14), 3)
'                End If
'            End If
'        End If
'    End If
'    If Val(GetSetting("HostForms_SA", "Settings", "SaveStringResult", "0")) = 1 Then
'        myConnection.Execute temps
'    Else
'        If Len(Trim(myResults2(5))) > 0 And IsNumeric(myResults2(5)) And Len(Trim(ResultOtherInfo(2))) > 0 Then myConnection.Execute temps
'    End If
'Set rstSearch = Nothing
'Next ctr1
'If Autoprint Then If Len(tPatientName) Then myConnection.Execute "Insert Into LabTempPrint (InstrumentID, SpecimenID, PatientName, Resultdate) select '" & InstrumentID & "','" & tSpecimenID & "','" & tPatientName & "','" & tDate & "'"
'myConnection.CommitTrans
'End If
'End Function

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
            'Roll back before showing the message so no partial result is left in an open transaction
            hasBeginTrans = False
            On Error Resume Next
            myConnection.RollbackTrans
            On Error GoTo 0
        End If
        Select Case errNum
            Case 9: MsgBox "Error in conversion. Unable to process data.", vbOKOnly + vbCritical, "Conversion Error"
            Case Else: MsgBox errNum & " " & errDesc, vbCritical + vbOKOnly, "Error"
        End Select
    End If
End Function

Private Sub LogComTrans(LogMsg As String, mIndex As Integer)
    Dim mComLogFile As Object
    Dim TempLogFile As String
    
    TempLogFile = Replace(LogMsg, cSOH(mIndex), "<SOH>")
    TempLogFile = Replace(TempLogFile, cSTX(mIndex), "<STX>")
    TempLogFile = Replace(TempLogFile, cETX(mIndex), "<ETX>")
    TempLogFile = Replace(TempLogFile, cEOT(mIndex), "<EOT>")
    TempLogFile = Replace(TempLogFile, cENQ(mIndex), "<ENQ>")
    TempLogFile = Replace(TempLogFile, cACK(mIndex), "<ACK>")
    TempLogFile = Replace(TempLogFile, cNAK(mIndex), "<NAK>")
    TempLogFile = Replace(TempLogFile, cSYN(mIndex), "<SYN>")
    TempLogFile = Replace(TempLogFile, Chr(13), "<CR>")
    TempLogFile = Replace(TempLogFile, Chr(10), "<LF>")
    
    TempLogFile = Format(Now, "YYYYmmddHHmmss") & " " & TempLogFile & vbCrLf
    If Len(Dir(App.Path & "\ComTransLog", vbDirectory)) = 0 Then
        fsoFiles.Createfolder App.Path & "\ComTransLog"
    End If
    'Set mComLogFile = fsoFiles.OpenTextFile(App.Path & "\ComTransLog\" & InstID(mIndex) & Format(Now, "YYYYmmdd") & ".txt", 8, True)
    Set mComLogFile = fsoFiles.OpenTextFile(App.Path & "\ComTransLog\" & InstrumentID & Format(Now, "YYYYmmdd") & ".txt", 8, True)
    mComLogFile.Write TempLogFile
    mComLogFile.Close
End Sub

Private Sub SendtoCom(mTrans As Integer, mStr As String, Optional mIndex As Integer = 0)
Dim tempstr As String
tempstr = ""
Select Case mTrans
    Case 1
        mscMain(mIndex).Output = Chr(CInt(mStr))
        Select Case mStr
            Case "1": DebugMode "<HOST(" & mIndex & ")>" & "SOH", mIndex
            Case "2": DebugMode "<HOST(" & mIndex & ")>" & "STX", mIndex
            Case "3": DebugMode "<HOST(" & mIndex & ")>" & "ETX", mIndex
            Case "4": DebugMode "<HOST(" & mIndex & ")>" & "EOT", mIndex
            Case "5": DebugMode "<HOST(" & mIndex & ")>" & "ENQ", mIndex
            Case "6": DebugMode "<HOST(" & mIndex & ")>" & "ACK", mIndex
            Case "15": DebugMode "<HOST(" & mIndex & ")>" & "NAK", mIndex
            Case "16": DebugMode "<HOST(" & mIndex & ")>" & "SYN", mIndex
        End Select
    Case 2
        mscMain(mIndex).Output = mStr
        If blnLogTrans Then
            LogComTrans "<LMS> " & mStr, mIndex
        End If
'        Debug.Print "<HOST>(" & mIndex & ")>" & mStr
'        Debug.Print mscMain(mIndex).OutBufferCount
End Select
'Debug.Print vbCrLf
'Debug.Print "----------------------------------"
'Debug.Print "               WAITEOT:" & waitEot(mIndex)
'Debug.Print "               WAITACK:" & WaitAck(mIndex)
'Debug.Print "               WAITSOH:" & waitSoh(mIndex)
'Debug.Print "               WAITENQ:" & WaitEnq(mIndex)
'Debug.Print "               ACKED:" & Acked(mIndex)
'Debug.Print "               RECEIVEMODE:" & ReceiveMode(mIndex)
'Debug.Print "               SENDMODE:" & SendMode(mIndex)
'Debug.Print "               CURRBATCHCOUNTER:" & CurrBatchCounter(mIndex)
'Debug.Print "               BATCHCOUNTER:" & BatchCounter(mIndex)
'Debug.Print "               REQUEST SENT:" & RequestSent(mIndex)
'Debug.Print "               WLTIMER:" & WLTimer(mIndex)
'Debug.Print "----------------------------------"
End Sub

Private Sub DebugMode(mReceived As String, mIndex As Integer)
If Val(GetSetting("ComSet", "Settings", "DebugMode", "0")) Then
    displayresults mReceived
    'displayresults "QR: " & myQueryRequests(mIndex)
    'displayresults "RM:" & Bool2Int(ReceiveMode(mIndex)) & " SM:" & Bool2Int(SendMode(mIndex)) & " AC:" & Bool2Int(Acked(mIndex)) & _
    '" WA:" & Bool2Int(WaitAck(mIndex)) & " WE:" & Bool2Int(waitEot(mIndex)) & " RS:" & Bool2Int(RequestSent(mIndex))
    displayresults "-------------------------"
End If
End Sub

