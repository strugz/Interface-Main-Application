VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMiddleware
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Middleware (LIS/HIS) Settings"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSWinsockLib.Winsock wsTest
      Left            =   120
      Top             =   6540
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrRefresh
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   600
      Top             =   6540
   End
   Begin VB.CommandButton cmdClose
      Caption         =   "&Close"
      Height          =   400
      Left            =   8160
      TabIndex        =   3
      Top             =   6540
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave
      Caption         =   "&Save"
      Height          =   400
      Left            =   6720
      TabIndex        =   2
      Top             =   6540
      Width           =   1320
   End
   Begin VB.OptionButton optTab
      Caption         =   "Outbox"
      Height          =   375
      Index           =   1
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1500
   End
   Begin VB.OptionButton optTab
      Caption         =   "Settings"
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1500
   End
   Begin VB.Frame fraOutbox
      BorderStyle     =   0  'None
      Height          =   5820
      Left            =   120
      TabIndex        =   40
      Top             =   600
      Visible         =   0   'False
      Width           =   9360
      Begin VB.TextBox txtView
         BeginProperty Font
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   49
         Top             =   4020
         Width           =   9360
      End
      Begin VB.CommandButton cmdView
         Caption         =   "&View"
         Height          =   375
         Left            =   3240
         TabIndex        =   48
         Top             =   3540
         Width           =   1320
      End
      Begin VB.CommandButton cmdResend
         Caption         =   "Re&send Selected"
         Height          =   375
         Left            =   1440
         TabIndex        =   47
         Top             =   3540
         Width           =   1680
      End
      Begin VB.CommandButton cmdRefresh
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   0
         TabIndex        =   46
         Top             =   3540
         Width           =   1320
      End
      Begin VB.ComboBox cboFolder
         Height          =   315
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   0
         Width           =   1500
      End
      Begin VB.ComboBox cboDest
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   0
         Width           =   1800
      End
      Begin MSComctlLib.ListView lvwOutbox
         Height          =   3000
         Left            =   0
         TabIndex        =   45
         Top             =   420
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   5292
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblHint
         Caption         =   "Double-click a message to see the HL7 that was built."
         Height          =   375
         Left            =   4680
         TabIndex        =   50
         Top             =   3600
         Width           =   4680
      End
      Begin VB.Label lblCounts
         Height          =   255
         Left            =   5280
         TabIndex        =   51
         Top             =   60
         Width           =   4080
      End
      Begin VB.Label Label20
         Caption         =   "Show:"
         Height          =   255
         Left            =   3000
         TabIndex        =   43
         Top             =   60
         Width           =   600
      End
      Begin VB.Label Label19
         Caption         =   "Destination:"
         Height          =   255
         Left            =   0
         TabIndex        =   41
         Top             =   60
         Width           =   1000
      End
   End
   Begin VB.Frame fraSettings
      BorderStyle     =   0  'None
      Height          =   5820
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   9360
      Begin VB.Frame fraDelivery
         Caption         =   "Delivery"
         Height          =   2280
         Left            =   4680
         TabIndex        =   31
         Top             =   3420
         Width           =   4680
         Begin VB.TextBox txtOutboxDir
            Height          =   285
            Left            =   1440
            TabIndex        =   38
            Top             =   1380
            Width           =   3120
         End
         Begin VB.TextBox txtKeepDays
            Height          =   285
            Left            =   2520
            TabIndex        =   36
            Top             =   1020
            Width           =   900
         End
         Begin VB.TextBox txtMaxAttempts
            Height          =   285
            Left            =   2520
            TabIndex        =   34
            Top             =   660
            Width           =   900
         End
         Begin VB.TextBox txtRetry
            Height          =   285
            Left            =   2520
            TabIndex        =   32
            Top             =   300
            Width           =   900
         End
         Begin VB.Label lblOutboxNote
            Caption         =   "Blank = default folder"
            Height          =   375
            Left            =   1440
            TabIndex        =   39
            Top             =   1740
            Width           =   3120
         End
         Begin VB.Label Label18
            Caption         =   "Outbox folder:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1410
            Width           =   1260
         End
         Begin VB.Label Label17
            Caption         =   "Keep sent messages (days):"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1050
            Width           =   2340
         End
         Begin VB.Label Label16
            Caption         =   "Give up after (attempts):"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   690
            Width           =   2340
         End
         Begin VB.Label Label15
            Caption         =   "First retry after (seconds):"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   330
            Width           =   2340
         End
      End
      Begin VB.Frame fraMessage
         Caption         =   "HL7 Message"
         Height          =   2280
         Left            =   0
         TabIndex        =   21
         Top             =   3420
         Width           =   4560
         Begin VB.CheckBox chkUnmapped
            Caption         =   "Send tests that have no LIS code (use analyzer code)"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1860
            Width           =   4320
         End
         Begin VB.TextBox txtRecvFac
            Height          =   285
            Left            =   2040
            TabIndex        =   28
            Top             =   1380
            Width           =   2400
         End
         Begin VB.TextBox txtRecvApp
            Height          =   285
            Left            =   2040
            TabIndex        =   26
            Top             =   1020
            Width           =   2400
         End
         Begin VB.TextBox txtSendFac
            Height          =   285
            Left            =   2040
            TabIndex        =   24
            Top             =   660
            Width           =   2400
         End
         Begin VB.TextBox txtSendApp
            Height          =   285
            Left            =   2040
            TabIndex        =   22
            Top             =   300
            Width           =   2400
         End
         Begin VB.Label Label14
            Caption         =   "Receiving facility:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1410
            Width           =   1860
         End
         Begin VB.Label Label13
            Caption         =   "Receiving application:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1050
            Width           =   1860
         End
         Begin VB.Label Label12
            Caption         =   "Sending facility:"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   690
            Width           =   1860
         End
         Begin VB.Label Label11
            Caption         =   "Sending application:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   330
            Width           =   1860
         End
      End
      Begin VB.Frame fraFile
         Caption         =   "HL7 Files to a Folder"
         Height          =   2400
         Left            =   4680
         TabIndex        =   14
         Top             =   900
         Width           =   4680
         Begin VB.CheckBox chkCRLF
            Caption         =   "End each line with CR+LF (off = CR only)"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1380
            Width           =   4440
         End
         Begin VB.TextBox txtFileExt
            Height          =   285
            Left            =   1080
            TabIndex        =   18
            Top             =   960
            Width           =   1000
         End
         Begin VB.TextBox txtFileDir
            Height          =   285
            Left            =   1080
            TabIndex        =   16
            Top             =   600
            Width           =   3480
         End
         Begin VB.CheckBox chkFile
            Caption         =   "Write each result as an HL7 file"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   280
            Width           =   4440
         End
         Begin VB.Label lblFileNote
            Caption         =   "The folder is the same one used by HL7 Settings (output folder)."
            Height          =   495
            Left            =   120
            TabIndex        =   52
            Top             =   1740
            Width           =   4440
         End
         Begin VB.Label Label10
            Caption         =   "Extension:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   990
            Width           =   900
         End
         Begin VB.Label Label9
            Caption         =   "Folder:"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   630
            Width           =   900
         End
      End
      Begin VB.Frame fraTcp
         Caption         =   "HL7 over TCP (MLLP) to the LIS/HIS"
         Height          =   2400
         Left            =   0
         TabIndex        =   7
         Top             =   900
         Width           =   4560
         Begin VB.CommandButton cmdTest
            Caption         =   "&Test Connection"
            Height          =   375
            Left            =   2760
            TabIndex        =   13
            Top             =   1620
            Width           =   1680
         End
         Begin VB.CheckBox chkKeepAlive
            Caption         =   "Keep connection open"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1680
            Width           =   2520
         End
         Begin VB.TextBox txtAckTimeout
            Height          =   285
            Left            =   1680
            TabIndex        =   11
            Top             =   1320
            Width           =   1000
         End
         Begin VB.TextBox txtPort
            Height          =   285
            Left            =   1680
            TabIndex        =   10
            Top             =   960
            Width           =   1000
         End
         Begin VB.TextBox txtHost
            Height          =   285
            Left            =   1680
            TabIndex        =   9
            Top             =   600
            Width           =   2760
         End
         Begin VB.CheckBox chkTcp
            Caption         =   "Send results over TCP"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   280
            Width           =   4320
         End
         Begin VB.Label lblTest
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   2060
            Width           =   4320
         End
         Begin VB.Label Label8
            Caption         =   "ACK timeout (s):"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   1350
            Width           =   1500
         End
         Begin VB.Label Label7
            Caption         =   "Port:"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   990
            Width           =   1500
         End
         Begin VB.Label Label6
            Caption         =   "LIS host / IP:"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   630
            Width           =   1500
         End
      End
      Begin VB.Frame fraMode
         Caption         =   "Mode"
         Height          =   780
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   9360
         Begin VB.CommandButton cmdSwitch
            Caption         =   "Switch to Middleware Mode..."
            Height          =   400
            Left            =   6600
            TabIndex        =   6
            Top             =   240
            Width           =   2640
         End
         Begin VB.Label lblMode
            Height          =   435
            Left            =   120
            TabIndex        =   58
            Top             =   270
            Width           =   6360
         End
      End
   End
End
Attribute VB_Name = "frmMiddleware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' Middleware (LIS/HIS) settings and Outbox monitor. See docs\MIDDLEWARE_MODE.md.
' Settings tab: registry Lab-I\Middleware (+ LAB-i\HL7 Settings for the shared keys).
'   modOutbox reads them on every use, so saved changes apply right away;
'   only the mode itself needs a restart (cmdSwitch -> mdiMain.SwitchModeWithRestart).
' Outbox tab: lists the spool files; it never touches the database.
'==============================================================================
Option Explicit

Private Const MAX_LIST As Long = 500

Private mSnapshot As String         'field values as loaded/saved, to detect unsaved edits
Private mTestState As Integer       '0 idle, 1 connecting, 2 connected, 3 failed
Private mTestError As String
Private mCloseChecked As Boolean

'------------------------------------------------------------------------------
' Form
'------------------------------------------------------------------------------
Private Sub Form_Load()
    With lvwOutbox.ColumnHeaders
        .Clear
        .Add , , "Time", 1650
        .Add , , "Sample No.", 1600
        .Add , , "Control ID", 1750
        .Add , , "Tries", 600
        .Add , , "Status", 1500
        .Add , , "Last result / error", 4000
    End With
    cboDest.AddItem "HL7 over TCP"
    cboDest.AddItem "HL7 files"
    If MW_TcpEnabled Or Not MW_FileEnabled Then cboDest.ListIndex = 0 Else cboDest.ListIndex = 1
    cboFolder.AddItem "Pending"
    cboFolder.AddItem "Failed"
    cboFolder.AddItem "Sent"
    cboFolder.ListIndex = 0
    LoadSettings
    ShowTab 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mCloseChecked Then Exit Sub
    If Not ConfirmUnsaved() Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    tmrRefresh.Enabled = False
    If wsTest.State <> sckClosed Then wsTest.Close
End Sub

'Returns False when the user cancels. Offers to save unsaved edits.
Private Function ConfirmUnsaved() As Boolean
    ConfirmUnsaved = True
    If FieldsSnapshot() = mSnapshot Then Exit Function
    Select Case MsgBox("Save the changes to the Middleware settings?", vbQuestion + vbYesNoCancel, "Middleware Settings")
        Case vbYes
            If ValidateSettings() Then
                SaveSettings
            Else
                ConfirmUnsaved = False
            End If
        Case vbNo
        Case Else
            ConfirmUnsaved = False
    End Select
End Function

Private Sub cmdClose_Click()
    If Not ConfirmUnsaved() Then Exit Sub
    mCloseChecked = True
    Unload Me
End Sub

Private Sub optTab_Click(Index As Integer)
    ShowTab Index
End Sub

Private Sub ShowTab(ByVal Index As Integer)
    fraSettings.Visible = (Index = 0)
    fraOutbox.Visible = (Index = 1)
    cmdSave.Visible = (Index = 0)
    tmrRefresh.Enabled = (Index = 1)
    If Index = 1 Then RefreshList
End Sub

'------------------------------------------------------------------------------
' Settings tab
'------------------------------------------------------------------------------
Private Sub LoadSettings()
    If MiddlewareMode Then
        lblMode.Caption = "Current mode: MIDDLEWARE. Results are forwarded to the LIS/HIS only; " & _
                          "search, result entry, reports and printing are hidden."
        cmdSwitch.Caption = "Switch to Main Mode..."
    Else
        lblMode.Caption = "Current mode: MAIN (normal Lab-I). The settings below are used " & _
                          "when Middleware mode is turned on."
        cmdSwitch.Caption = "Switch to Middleware Mode..."
    End If

    chkTcp.Value = IIf(MW_TcpEnabled, vbChecked, vbUnchecked)
    txtHost.Text = MW_Setting("HL7TcpHost", "")
    txtPort.Text = MW_Setting("HL7TcpPort", "")
    txtAckTimeout.Text = MW_Setting("AckTimeout", "10")
    chkKeepAlive.Value = IIf(Val(MW_Setting("KeepAlive", "1")) = 1, vbChecked, vbUnchecked)

    chkFile.Value = IIf(MW_FileEnabled, vbChecked, vbUnchecked)
    txtFileDir.Text = GetSetting("LAB-i", "HL7 Settings", "FWDIR", App.Path & "\HostSpooler")
    txtFileExt.Text = MW_Setting("HL7FileExt", "hl7")
    chkCRLF.Value = IIf(Val(MW_Setting("HL7FileCRLF", "1")) = 1, vbChecked, vbUnchecked)

    txtSendApp.Text = GetSetting("LAB-i", "HL7 Settings", "Sending Application", "LAB-Iv" & App.Major)
    txtSendFac.Text = MW_Setting("SendingFacility", "")
    txtRecvApp.Text = GetSetting("LAB-i", "HL7 Settings", "Receiving Application", "MDMPI")
    txtRecvFac.Text = MW_Setting("ReceivingFacility", "")
    chkUnmapped.Value = IIf(Val(MW_Setting("SendUnmapped", "1")) = 1, vbChecked, vbUnchecked)

    txtRetry.Text = MW_Setting("RetrySeconds", "30")
    txtMaxAttempts.Text = MW_Setting("MaxAttempts", "10")
    txtKeepDays.Text = MW_Setting("KeepSentDays", "90")
    txtOutboxDir.Text = MW_Setting("OutboxDir", "")
    lblOutboxNote.Caption = "Blank = default: " & DefaultOutboxText()

    lblTest.Caption = ""
    mSnapshot = FieldsSnapshot()
End Sub

Private Function DefaultOutboxText() As String
    If DBServerType = "LOCAL DATABASE" Then
        DefaultOutboxText = "data folder\Outbox"
    Else
        DefaultOutboxText = "program folder\Outbox"
    End If
End Function

Private Function FieldsSnapshot() As String
    FieldsSnapshot = chkTcp.Value & "|" & Trim(txtHost.Text) & "|" & Trim(txtPort.Text) & "|" & Trim(txtAckTimeout.Text) & "|" & _
        chkKeepAlive.Value & "|" & chkFile.Value & "|" & Trim(txtFileDir.Text) & "|" & Trim(txtFileExt.Text) & "|" & chkCRLF.Value & "|" & _
        Trim(txtSendApp.Text) & "|" & Trim(txtSendFac.Text) & "|" & Trim(txtRecvApp.Text) & "|" & Trim(txtRecvFac.Text) & "|" & _
        chkUnmapped.Value & "|" & Trim(txtRetry.Text) & "|" & Trim(txtMaxAttempts.Text) & "|" & Trim(txtKeepDays.Text) & "|" & Trim(txtOutboxDir.Text)
End Function

Private Function IsWholeNumber(ByVal s As String, ByVal MinValue As Long, ByVal MaxValue As Long) As Boolean
    s = Trim(s)
    If Len(s) = 0 Or Len(s) > 9 Then Exit Function
    If s Like "*[!0-9]*" Then Exit Function
    IsWholeNumber = (CLng(s) >= MinValue And CLng(s) <= MaxValue)
End Function

Private Function Fail(ctl As Control, ByVal Message As String) As Boolean
    optTab(0).Value = True
    MsgBox Message, vbExclamation + vbOKOnly, "Middleware Settings"
    On Error Resume Next
    ctl.SetFocus
    Fail = False
End Function

Private Function ValidateSettings() As Boolean
    ValidateSettings = False
    If chkTcp.Value = vbChecked Then
        If Len(Trim(txtHost.Text)) = 0 Then
            Fail txtHost, "Enter the LIS host name or IP address."
            Exit Function
        End If
        If Not IsWholeNumber(txtPort.Text, 1, 65535) Then
            Fail txtPort, "Enter the LIS port number (1 to 65535)."
            Exit Function
        End If
    End If
    If Not IsWholeNumber(txtAckTimeout.Text, 2, 300) Then
        Fail txtAckTimeout, "ACK timeout must be 2 to 300 seconds."
        Exit Function
    End If
    If chkFile.Value = vbChecked Then
        If Len(Trim(txtFileDir.Text)) = 0 Then
            Fail txtFileDir, "Enter the folder where the HL7 files are written."
            Exit Function
        End If
        If Not fsoFiles.FolderExists(Trim(txtFileDir.Text)) Then
            If MsgBox("The folder " & Trim(txtFileDir.Text) & " does not exist (or cannot be reached)." & vbCrLf & _
                      "Messages will wait in the Outbox until it can be reached. Save anyway?", vbQuestion + vbYesNo, "Middleware Settings") <> vbYes Then
                Fail txtFileDir, "Check the folder."
                Exit Function
            End If
        End If
    End If
    If chkTcp.Value <> vbChecked And chkFile.Value <> vbChecked Then
        If MsgBox("No destination is turned on, so results will not be sent anywhere in Middleware mode. Save anyway?", _
                  vbQuestion + vbYesNo, "Middleware Settings") <> vbYes Then Exit Function
    End If
    If Not IsWholeNumber(txtRetry.Text, 5, 3600) Then
        Fail txtRetry, "First retry must be 5 to 3600 seconds."
        Exit Function
    End If
    If Not IsWholeNumber(txtMaxAttempts.Text, 1, 1000) Then
        Fail txtMaxAttempts, "Attempts must be 1 to 1000."
        Exit Function
    End If
    If Not IsWholeNumber(txtKeepDays.Text, 0, 3650) Then
        Fail txtKeepDays, "Days must be 0 to 3650 (0 = keep forever)."
        Exit Function
    End If
    If InStr(1, txtFileExt.Text, "\") > 0 Or InStr(1, txtFileExt.Text, ".") > 0 Then
        Fail txtFileExt, "Enter the extension without a dot, e.g. hl7"
        Exit Function
    End If
    ValidateSettings = True
End Function

Private Sub SaveSettings()
    SaveSetting "Lab-I", "Middleware", "HL7Tcp", IIf(chkTcp.Value = vbChecked, "1", "0")
    SaveSetting "Lab-I", "Middleware", "HL7TcpHost", Trim(txtHost.Text)
    SaveSetting "Lab-I", "Middleware", "HL7TcpPort", Trim(txtPort.Text)
    SaveSetting "Lab-I", "Middleware", "AckTimeout", Trim(txtAckTimeout.Text)
    SaveSetting "Lab-I", "Middleware", "KeepAlive", IIf(chkKeepAlive.Value = vbChecked, "1", "0")

    SaveSetting "Lab-I", "Middleware", "HL7File", IIf(chkFile.Value = vbChecked, "1", "0")
    If Len(Trim(txtFileDir.Text)) > 0 Then SaveSetting "LAB-i", "HL7 Settings", "FWDIR", Trim(txtFileDir.Text)
    SaveSetting "Lab-I", "Middleware", "HL7FileExt", Trim(txtFileExt.Text)
    SaveSetting "Lab-I", "Middleware", "HL7FileCRLF", IIf(chkCRLF.Value = vbChecked, "1", "0")

    SaveSetting "LAB-i", "HL7 Settings", "Sending Application", Trim(txtSendApp.Text)
    SaveSetting "Lab-I", "Middleware", "SendingFacility", Trim(txtSendFac.Text)
    SaveSetting "LAB-i", "HL7 Settings", "Receiving Application", Trim(txtRecvApp.Text)
    SaveSetting "Lab-I", "Middleware", "ReceivingFacility", Trim(txtRecvFac.Text)
    SaveSetting "Lab-I", "Middleware", "SendUnmapped", IIf(chkUnmapped.Value = vbChecked, "1", "0")

    SaveSetting "Lab-I", "Middleware", "RetrySeconds", Trim(txtRetry.Text)
    SaveSetting "Lab-I", "Middleware", "MaxAttempts", Trim(txtMaxAttempts.Text)
    SaveSetting "Lab-I", "Middleware", "KeepSentDays", Trim(txtKeepDays.Text)
    SaveSetting "Lab-I", "Middleware", "OutboxDir", Trim(txtOutboxDir.Text)
    mSnapshot = FieldsSnapshot()
End Sub

Private Sub cmdSave_Click()
    If Not ValidateSettings() Then Exit Sub
    SaveSettings
    MsgBox "Middleware settings saved." & IIf(MiddlewareMode, " They apply to the next message.", ""), vbInformation + vbOKOnly, "Middleware Settings"
End Sub

Private Sub cmdSwitch_Click()
    If Not ConfirmUnsaved() Then Exit Sub
    If Not MiddlewareMode And chkTcp.Value <> vbChecked And chkFile.Value <> vbChecked Then
        If MsgBox("No destination is turned on yet, so results would not be sent anywhere. Switch anyway?", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "Middleware Settings") <> vbYes Then Exit Sub
    End If
    MW_SwitchRequested = True      'the caller restarts Lab-I in the other mode after this screen closes
    mCloseChecked = True
    Unload Me
End Sub

'Checks only that the LIS accepts a TCP connection on host:port. No HL7 message is
'sent, so nothing can end up in the LIS as a test result.
Private Sub cmdTest_Click()
Dim mStart As Single
Dim mTimeout As Long
On Error GoTo errFound
    If Len(Trim(txtHost.Text)) = 0 Then
        Fail txtHost, "Enter the LIS host name or IP address."
        Exit Sub
    End If
    If Not IsWholeNumber(txtPort.Text, 1, 65535) Then
        Fail txtPort, "Enter the LIS port number (1 to 65535)."
        Exit Sub
    End If
    If IsWholeNumber(txtAckTimeout.Text, 2, 300) Then mTimeout = CLng(txtAckTimeout.Text) Else mTimeout = 10
    cmdTest.Enabled = False
    lblTest.ForeColor = vbButtonText
    lblTest.Caption = "Connecting to " & Trim(txtHost.Text) & ":" & Trim(txtPort.Text) & "..."
    If wsTest.State <> sckClosed Then wsTest.Close
    wsTest.Protocol = sckTCPProtocol
    wsTest.RemoteHost = Trim(txtHost.Text)
    wsTest.RemotePort = CLng(txtPort.Text)
    wsTest.LocalPort = 0
    mTestState = 1
    mTestError = ""
    wsTest.Connect
    mStart = Timer
    Do While mTestState = 1
        DoEvents
        If Abs(Timer - mStart) > mTimeout Then
            mTestState = 3
            mTestError = "no answer within " & mTimeout & " s"
        End If
    Loop
    If wsTest.State <> sckClosed Then wsTest.Close
    If mTestState = 2 Then
        lblTest.ForeColor = &H8000&
        lblTest.Caption = "OK: the LIS is listening (connection test only, nothing sent)."
    Else
        lblTest.ForeColor = vbRed
        lblTest.Caption = "Failed: " & mTestError
    End If
    mTestState = 0
    cmdTest.Enabled = True
    Exit Sub
errFound:
    mTestState = 0
    lblTest.ForeColor = vbRed
    lblTest.Caption = "Failed: " & Err.Description
    cmdTest.Enabled = True
End Sub

Private Sub wsTest_Connect()
    If mTestState = 1 Then mTestState = 2
End Sub

Private Sub wsTest_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, _
                         ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CancelDisplay = True
    If mTestState = 1 Then
        mTestError = Description
        mTestState = 3
    End If
End Sub

'------------------------------------------------------------------------------
' Outbox tab
'------------------------------------------------------------------------------
Private Function SelectedDest() As String
    If cboDest.ListIndex = 1 Then SelectedDest = MW_DEST_FILE Else SelectedDest = MW_DEST_TCP
End Function

Private Function SelectedFolder() As String
    Select Case cboFolder.ListIndex
        Case 1: SelectedFolder = "failed"
        Case 2: SelectedFolder = "sent"
        Case Else: SelectedFolder = "pending"
    End Select
End Function

Private Sub cboDest_Click()
    If fraOutbox.Visible Then RefreshList
End Sub

Private Sub cboFolder_Click()
    If fraOutbox.Visible Then RefreshList
End Sub

Private Sub cmdRefresh_Click()
    RefreshList
End Sub

Private Sub tmrRefresh_Timer()
    UpdateCounts -1
End Sub

Private Sub UpdateCounts(ByVal Shown As Long)
Dim mDest As String
Dim s As String
    mDest = SelectedDest()
    s = "Pending " & OutboxCount(mDest, "pending") & "   Failed " & OutboxCount(mDest, "failed") & "   Sent " & OutboxCount(mDest, "sent")
    If Shown >= 0 Then lblCounts.Tag = IIf(Shown >= MAX_LIST, "   (newest " & MAX_LIST & " shown)", "")
    lblCounts.Caption = s & lblCounts.Tag
End Sub

Private Function FormatControlTime(ByVal ControlID As String) As String
    'control ID = yymmddhhnnss + 4-digit sequence
    If Len(ControlID) >= 12 And IsNumeric(Left(ControlID, 12)) Then
        FormatControlTime = "20" & Mid(ControlID, 1, 2) & "-" & Mid(ControlID, 3, 2) & "-" & Mid(ControlID, 5, 2) & " " & _
                            Mid(ControlID, 7, 2) & ":" & Mid(ControlID, 9, 2) & ":" & Mid(ControlID, 11, 2)
    Else
        FormatControlTime = ""
    End If
End Function

Private Sub RefreshList()
Dim col As Collection
Dim mFile As Variant
Dim mTotal As Long
Dim Header() As String
Dim Payload As String
Dim itm As MSComctlLib.ListItem
Dim mFolder As String
Dim mStatus As String
On Error GoTo errFound
    Screen.MousePointer = vbHourglass
    mFolder = SelectedFolder()
    lvwOutbox.ListItems.Clear
    txtView.Text = ""
    Set col = OutboxListFiles(SelectedDest(), mFolder, MAX_LIST, mTotal)
    For Each mFile In col
        If OutboxReadSpool(CStr(mFile), Header, Payload) Then
            Set itm = lvwOutbox.ListItems.Add(, , FormatControlTime(Header(4)))
            itm.Tag = CStr(mFile)
            itm.SubItems(1) = Header(3)
            itm.SubItems(2) = Header(4)
            itm.SubItems(3) = Header(5)
            Select Case mFolder
                Case "pending"
                    If Len(Header(6)) = 14 Then
                        mStatus = "Retry at " & Mid(Header(6), 9, 2) & ":" & Mid(Header(6), 11, 2) & ":" & Mid(Header(6), 13, 2)
                    Else
                        mStatus = "Waiting to send"
                    End If
                Case "failed"
                    mStatus = "FAILED"
                Case Else
                    mStatus = "Sent"
            End Select
            itm.SubItems(4) = mStatus
            itm.SubItems(5) = Header(7)
        End If
    Next mFile
    UpdateCounts mTotal
    cmdResend.Caption = IIf(mFolder = "pending", "Re&try Now", "Re&send Selected")
    Screen.MousePointer = vbDefault
    Exit Sub
errFound:
    Screen.MousePointer = vbDefault
    lblCounts.Caption = "Could not read the Outbox: " & Err.Description
End Sub

Private Sub cmdResend_Click()
Dim itm As MSComctlLib.ListItem
Dim mCount As Long
Dim mDone As Long
Dim mPrompt As String
    For Each itm In lvwOutbox.ListItems
        If itm.Selected Then mCount = mCount + 1
    Next itm
    If mCount = 0 Then
        MsgBox "Select one or more messages first.", vbInformation + vbOKOnly, "Outbox"
        Exit Sub
    End If
    Select Case SelectedFolder()
        Case "sent"
            mPrompt = "Send " & mCount & " message(s) again?" & vbCrLf & vbCrLf & _
                      "The LIS/HIS already accepted them, so it may receive duplicate results."
        Case "failed"
            mPrompt = "Put " & mCount & " failed message(s) back in the queue?"
        Case Else
            mPrompt = "Try " & mCount & " waiting message(s) again now?"
    End Select
    If MsgBox(mPrompt, vbQuestion + vbYesNo + vbDefaultButton2, "Outbox") <> vbYes Then Exit Sub
    For Each itm In lvwOutbox.ListItems
        If itm.Selected Then
            If OutboxResendFile(itm.Tag, SelectedDest()) Then mDone = mDone + 1
        End If
    Next itm
    RefreshList
    MsgBox mDone & " of " & mCount & " message(s) queued.", vbInformation + vbOKOnly, "Outbox"
End Sub

Private Sub cmdView_Click()
Dim Header() As String
Dim Payload As String
    If lvwOutbox.SelectedItem Is Nothing Then
        MsgBox "Select a message first.", vbInformation + vbOKOnly, "Outbox"
        Exit Sub
    End If
    If OutboxReadSpool(lvwOutbox.SelectedItem.Tag, Header, Payload) Then
        Payload = Replace(Replace(Payload, vbCrLf, vbCr), vbCr, vbCrLf)
        txtView.Text = "Sample " & Header(3) & "   Control ID " & Header(4) & "   Tries " & Header(5) & vbCrLf & _
                       "Last: " & Header(7) & vbCrLf & String(60, "-") & vbCrLf & Payload
    Else
        txtView.Text = "The message is no longer in this folder (it may have just been sent). Press Refresh."
    End If
End Sub

Private Sub lvwOutbox_DblClick()
    cmdView_Click
End Sub
