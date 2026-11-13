VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H80000003&
   Caption         =   "Host Interface"
   ClientHeight    =   8085
   ClientLeft      =   -75
   ClientTop       =   360
   ClientWidth     =   14895
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   1110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2982
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":385C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4736
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5610
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":64EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":73C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":829E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7710
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10345
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "09/04/2025"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "4:11 PM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   1080
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   1905
      ButtonWidth     =   2963
      ButtonHeight    =   1852
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "          Connect          "
            Key             =   "Connect"
            Object.ToolTipText     =   "Connect"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "c1"
                  Text            =   "Connect(Receive Only)"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "c2"
                  Text            =   "Connect(Send/Receive)"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "c3"
                  Text            =   "Connect(Send ID)"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Disconnect"
            Key             =   "Disconnect"
            Object.ToolTipText     =   "Disconnect"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "Search"
            Object.ToolTipText     =   "Search"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Lab Exams"
            Key             =   "Lab Exams"
            Object.ToolTipText     =   "Laboratory Exams"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Communication"
            Key             =   "Communication"
            Object.ToolTipText     =   "Communication Settings"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "medtechs"
            Object.ToolTipText     =   "Add Medtech"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "doctors"
            Object.ToolTipText     =   "Add Doctors"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "pathologists"
            Object.ToolTipText     =   "Add Pathologists"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Database Manager"
            Key             =   "database"
            Object.ToolTipText     =   "Database Maintenance"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit Application"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Timer tmrDailyBackup 
         Interval        =   1000
         Left            =   10170
         Top             =   210
      End
      Begin MSWinsockLib.Winsock wsBroadcast 
         Left            =   6690
         Top             =   300
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSearch 
         Caption         =   "Search Result"
         Checked         =   -1  'True
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuStat 
         Caption         =   "Statistics"
      End
      Begin VB.Menu mnuRequest 
         Caption         =   "Requests"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuWorklist 
         Caption         =   "Worklist"
         Shortcut        =   ^W
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQC 
         Caption         =   "Quality Control"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuConnection 
      Caption         =   "&Connection"
      Begin VB.Menu mnuConnect 
         Caption         =   "Connect(Receive Only)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuSendReceive 
         Caption         =   "Connect(Send/Receive)"
         Shortcut        =   ^R
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSendID 
         Caption         =   "Connect(Send ID)"
         Shortcut        =   ^I
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep00 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
         Shortcut        =   ^D
      End
      Begin VB.Menu mnusep18 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRemote 
         Caption         =   "Connect to Remote Computer"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuBuildFile 
      Caption         =   "Maintenance"
      Visible         =   0   'False
      Begin VB.Menu mnuLabExam 
         Caption         =   "&Exams"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnusep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClass 
         Caption         =   "N.V. Classi&fications"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuNvalues 
         Caption         =   "Normal &Values"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnusep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLabprofiles 
         Caption         =   "Profiles"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuprofiles 
         Caption         =   "Profi&le Exams"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnusep09 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHospital 
         Caption         =   "&Hospital Profile"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuPathologist 
         Caption         =   "Pathologist"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuDoctors 
         Caption         =   "Physician"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuMedtech 
         Caption         =   "Med. Tech."
         Shortcut        =   ^M
      End
      Begin VB.Menu sep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFlagLists 
         Caption         =   "Flag List&s"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSelectionLists 
         Caption         =   "Selection Lists &Update..."
         Shortcut        =   ^U
      End
      Begin VB.Menu mnusep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportFormats 
         Caption         =   "Repor&t Formats"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuWorklistReport 
         Caption         =   "Worklist Report"
      End
      Begin VB.Menu mnuLogbook 
         Caption         =   "Summary Logbook"
      End
      Begin VB.Menu mnuOtherReports 
         Caption         =   "Other Reports..."
      End
      Begin VB.Menu mnusep08 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuReprint 
         Caption         =   "Reprint Result"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuControlPanel 
         Caption         =   "Pr&ogram Setup..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuProgSettings 
         Caption         =   "Program Settings"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFormSet 
         Caption         =   "Result Form Settings"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep03 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuComSet 
         Caption         =   "Communication Settings"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuConSet 
         Caption         =   "Database Connection Settings"
      End
      Begin VB.Menu mnuConnServ 
         Caption         =   "Connect to Server..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInitDTP
         Caption         =   "Initialize DTP..."
      End
      Begin VB.Menu mnuSep06
         Caption         =   "-"
      End
      Begin VB.Menu mnuMiddleware
         Caption         =   "Middleware Settings..."
      End
      Begin VB.Menu mnuSwitchMode
         Caption         =   "Switch to Middleware Mode..."
      End
      Begin VB.Menu mnuSep05 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options..."
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuBarCode 
         Caption         =   "Print Barcode"
      End
      Begin VB.Menu mnuFixAttachment 
         Caption         =   "Fix Database Attachments..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuClearBuffer 
         Caption         =   "Clear Database Buffer..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuArchive 
         Caption         =   "Archive Database"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBackUp 
         Caption         =   "Database Maintenance..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPlugTest 
         Caption         =   "Plugin Test..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuStatistics 
         Caption         =   "Statistics..."
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuUtilities 
      Caption         =   "&Utilities"
      Visible         =   0   'False
      Begin VB.Menu mnuConsole 
         Caption         =   "Command Console"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Visible         =   0   'False
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About Lab-I"
      End
      Begin VB.Menu mnuSep07 
         Caption         =   "-"
      End
      Begin VB.Menu mnuactivation 
         Caption         =   "License Manager"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSkipExitPrompt As Boolean      'set by mnuSwitchMode_Click before the automatic restart

Private Sub MDIForm_Load()
Dim rstVersion As ADODB.Recordset
Dim isCorrectVersion As Boolean
Dim DInstrumentId As String
On Error GoTo errFound
'    Me.WindowState = 0
'    Me.Width = Screen.Width
'    Me.Height = Screen.Height
'    Me.Left = 1
'    Me.Top = 1
    InitCommonControlsVB
'    OpenWSBroadcast
    isCorrectVersion = True
    Set rstVersion = myConnection.Execute("Select * from tbVersion")
    If isCorrectVersion Then
        If rstVersion.Fields("Major").Value = "2" And Val(rstVersion.Fields("Minor").Value) < 55 Then
            sbrStatus.Panels(1).Text = "DB Version " & rstVersion.Fields("Major").Value & "." & rstVersion.Fields("Minor").Value
        Else
            sbrStatus.Panels(1).Text = "DB Version " & rstVersion.Fields("Major").Value & "." & rstVersion.Fields("Minor").Value & " (" & rstVersion.Fields("Date").Value & ")"
        End If
        Set rstVersion = Nothing
    End If
    sbrStatus.Panels(1).AutoSize = sbrContents
    sbrStatus.Panels(2).AutoSize = sbrSpring
    sbrStatus.Panels(5).Visible = False
    sbrStatus.Panels(6).Style = sbrText
    sbrStatus.Panels(6).Text = WinVersion
    sbrStatus.Panels(7).Visible = False
    mnuactivation.Visible = True
    If MiddlewareMode Then ApplyMiddlewareMenus
    GetDBSettings
    Set mComSet = CreateObject("Comset.ComSettings")
    If isDBconnected Then
        setHostForms
        If GetSetting("Lab-I", "Database", "ProcessError", "No") = "Yes" Then
            If MsgBox("Error has been logged during instrument result processing. Do you wish to clear the result buffer?", vbQuestion + vbYesNo + vbDefaultButton2, "Result Buffer Error") = vbYes Then
                myConnection.Execute "delete from LabUnprocessed where instrumentid = '" & InstrumentID & "'"
                MsgBox "Instrument Result Buffer has been cleared.", vbOKOnly + vbInformation, App.ProductName
            End If
            SaveSetting "Lab-I", "Database", "ProcessError", "No"
        End If
    End If
    sbrStatus.Panels(1).Text = sbrStatus.Panels(1).Text & "/SA." & HostSA.Version
    If InstrumentID = "HMX" Then DInstrumentId = "STKS/HMX/LH" Else DInstrumentId = InstrumentID
    Me.Caption = App.Title & "." & App.Minor & " - [rev. " & Str(App.Revision) & "] - [" & DInstrumentId & "]"
    If MiddlewareMode Then
        Me.Caption = Me.Caption & " - [MIDDLEWARE]"
        mnuSwitchMode.Caption = "Switch to Main Mode..."
    Else
        mnuSwitchMode.Caption = "Switch to Middleware Mode..."
    End If
    setButtons 1, 0, 1, 1, 1
    If EnableDTP Then
        If Val(ObjDTP.ErrorNum) <> 0 Then
            mnuInitDTP.Visible = True
        Else
            mnuInitDTP.Visible = False
        End If
    Else
        mnuInitDTP.Visible = False
    End If
    If Val(GetSetting("Lab-I", "Settings", "AutoConnect", "0")) = 1 Then
        mnuConnect_Click
        MDIForm_Resize
    End If
    BackupSched = GetSetting("Lab-I", "Database", "DailyBackup", "")
errFound:
    If Err.Number <> 0 Then
        Select Case Err.Number
            Case 429
                MsgBox "Failed to initialize program. Some components are not installed properly.", vbOKOnly + vbExclamation, App.ProductName
            Case 53
                Resume Next
            Case -2147217865
                sbrStatus.Panels(1).Text = "DB Version 1.0"
                isCorrectVersion = False
                Resume Next
            Case 3704
                mnuConSet_Click
            Case Else
                MsgBox Err.Description, vbOKOnly + vbExclamation, App.ProductName
        End Select
        End
    End If
End Sub

'Middleware mode (Lab-I\Middleware\Enable): hide result search, verification, printing,
'reports and the people/report master data. Exams, normal values and profiles stay,
'because they hold the LIS test code mapping (tbcolabexam.HisExamCode).
Private Sub ApplyMiddlewareMenus()
On Error Resume Next
    mnuSearch.Checked = False
    mnuSearch.Enabled = False
    mnuSearch.Visible = False
    mnuStat.Enabled = False
    mnuStat.Visible = False
    mnuRequest.Enabled = False
    mnuRequest.Visible = False
    mnuWorklist.Enabled = False
    mnuWorklist.Visible = False
    mnuQC.Enabled = False
    mnuQC.Visible = False
    mnuSep02.Visible = False
    mnuReports.Visible = False
    mnuHospital.Visible = False
    mnuPathologist.Visible = False
    mnuDoctors.Visible = False
    mnuMedtech.Visible = False
    sep01.Visible = False
    mnuFlagLists.Visible = False
    mnuSelectionLists.Visible = False
    mnusep12.Visible = False
    mnuReportFormats.Visible = False
    mnuBarCode.Visible = False
    mnuStatistics.Visible = False
    tbrMain.Buttons("Search").Value = tbrUnpressed
    tbrMain.Buttons("Search").Visible = False
    tbrMain.Buttons("medtechs").Visible = False
    tbrMain.Buttons("doctors").Visible = False
    tbrMain.Buttons("pathologists").Visible = False
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
    If Screen.Width <> initScreenSize Then
        Me.Top = 1
        Me.Left = 1
        Me.Width = Screen.Width
        Me.Height = Screen.Height
        initScreenSize = Screen.Width
'        DoEvents
'        If Screen.Width = 12000 Then
'            Me.Picture = LoadPicture(App.Path & "\" & "lab-i.sys")
'        Else
'            Me.Picture = LoadPicture(App.Path & "\" & "lab-i_L.sys")
'        End If
    End If

    If GetSetting("HostForms", "Forms", "Viewsearch", "0") = "1" And Not MiddlewareMode Then
        If EnableDTP And Val(GetSetting("Lab-I", "DaTran", "DisableSearch", "0")) = 1 Then
            mnuSearch.Checked = False
            tbrMain.Buttons("Search").Value = tbrUnpressed
            tbrMain.Buttons("Search").Enabled = False
            mnuSearch.Enabled = False
            frmSearch.Hide
            frmReceiveData.Top = 1
            frmReceiveData.Height = mdiMain.ScaleHeight
        Else
            mnuSearch.Checked = True
            tbrMain.Buttons("Search").Value = tbrPressed
            tbrMain.Buttons("Search").Enabled = True
            mnuSearch.Enabled = True
            frmSearch.Show
            If isConnected Then
'                If mdiMain.Height >= frmSearch.fraBody.Height + 50 + frmSearch.fraTitle.Height + frmReceiveData.Height Then
                    frmReceiveData.Top = 1
                    frmReceiveData.Height = mdiMain.ScaleHeight / 2
                    frmSearch.Top = frmReceiveData.Height
                    frmSearch.Height = mdiMain.ScaleHeight / 2 ' - frmReceiveData.Height 'mdiMain.Height - 800 - frmReceiveData.Height
'                End If
            Else
'                If mdiMain.Height >= frmSearch.fraBody.Height + 50 + frmSearch.fraTitle.Height Then
                    frmSearch.Top = 1
                    frmSearch.Height = mdiMain.ScaleHeight  'mdiMain.Height - 800
                    frmSearch.fraBody(0).Height = frmSearch.Height - 300
'                End If
            End If
        End If
    Else
        If isConnected Then
            frmReceiveData.Top = 1
            frmReceiveData.Height = mdiMain.ScaleHeight
        End If
        mnuSearch.Checked = False
        Unload frmSearch
    End If
    If isResultVisible Then
        frmSearchResult.Top = 1
        frmSearchResult.Height = mdiMain.ScaleHeight 'mdiMain.Height - 800
        frmSearchResult.Width = mdiMain.ScaleWidth - frmSearch.Width 'mdiMain.Width - 5100
        If mdiMain.mnuSearch.Checked Or isConnected Then
            frmSearchResult.Left = frmSearch.Width  '4815 + 80
        Else
            frmSearchResult.Left = 1
        End If
    End If
    frmWall.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim mExit As Boolean
If mSkipExitPrompt Then
    mExit = True                'restarting after a mode switch: already confirmed
Else
    mExit = (MsgBox("Are you sure you want to exit from the program?", vbYesNo + vbQuestion, App.ProductName) = vbYes)
End If
If mExit Then
    Unload frmProcess
    If isDBconnected Then myConnection.Close
    Set HostSA = Nothing
    Set myConnection = Nothing
Else
    Cancel = 1
    If mnuSearch.Checked Then frmSearch.Show
    If isConnected Then frmReceiveData.Show
    If isResultVisible Then frmSearchResult.Show
    frmWall.Show
End If
End Sub

Private Sub mnuAbout_Click()
'    frmAbout.Show 1
'frmDailyBackup.Show 1
End Sub

Private Sub mnuactivation_Click()
    frmLicenseManager.Show 1
    If Not IsRegistered Then
    
    End If
End Sub

Private Sub mnuBackUp_Click()
    If isResultVisible Then frmSearchResult.Visible = False
    frmDbMan.Show 1
End Sub

Private Sub mnuBarCode_Click()
    frmBarCode.Show 1
End Sub

Private Sub mnuClass_Click()
    HostSA.showClassifications
End Sub

Private Sub mnuClearBuffer_Click()
If MsgBox("Are you sure you want to clear the instrument database buffer?", vbCritical + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
    myConnection.Execute "delete from LabUnprocessed"
    MsgBox "Instrument Database Buffer has been cleared.", vbOKOnly + vbInformation, App.ProductName
End If
End Sub

Private Sub mnuComSet_Click()
    mComSet.showsettings
End Sub

Private Sub mnuConnect_Click()
On Error Resume Next
'    myCommStat = cs_ReceiveOnly
    Load frmReceiveData
    If isConnected Then setButtons 0, 1, 0, 0, 0 Else setButtons 1, 0, 1, 1, 0
End Sub

Private Sub mnuConnServ_Click()
OpenConnection
Me.mnuConnServ.Visible = Not isDBconnected
End Sub

Private Sub mnuConSet_Click()
Unload frmSearchResult
isResultVisible = False
frmDBsettings.Show 1
CloseConnection
GetDBSettings
OpenConnection
mdiMain.mnuConnServ.Visible = Not isDBconnected
End Sub

'Settings > Switch to Middleware Mode / Switch to Main Mode.
'Saves the mode (and the Main mode Autoprocess/Autoprint, see modMain.SetMiddlewareMode),
'then restarts Interface.exe so every menu and setting is applied cleanly.
Private Sub mnuSwitchMode_Click()
    SwitchModeWithRestart
End Sub

'Settings > Middleware Settings... (also Program Setup > 17 - Middleware).
Private Sub mnuMiddleware_Click()
    ShowMiddlewareSettings
End Sub

Public Sub ShowMiddlewareSettings()
    MW_SwitchRequested = False
    frmMiddleware.Show 1
    If MW_SwitchRequested Then
        MW_SwitchRequested = False
        SwitchModeWithRestart
    End If
End Sub

Public Sub SwitchModeWithRestart()
Dim mPrompt As String
Dim mExe As String
On Error GoTo errFound
    If isConnected Then
        If MsgBox("The instrument must be disconnected before changing mode. Disconnect now?", vbQuestion + vbYesNo + vbDefaultButton2, "Change Mode") = vbYes Then
            mnuDisconnect_Click
        End If
        If isConnected Then Exit Sub
    End If
    If MiddlewareMode Then
        mPrompt = "Switch back to MAIN mode?" & vbCrLf & vbCrLf & _
                  "Search, result entry, reports and printing come back, and the Autoprocess/Autoprint settings " & _
                  "you had before Middleware mode are restored." & vbCrLf & _
                  "Results already in the Outbox stay there and are sent the next time Middleware mode is on."
    Else
        mPrompt = "Switch to MIDDLEWARE mode?" & vbCrLf & vbCrLf & _
                  "Results are only forwarded to the LIS/HIS (Lab-I\Middleware settings). Search, result entry, " & _
                  "reports and printing are hidden. Your current Autoprocess/Autoprint settings are saved and " & _
                  "restored when you switch back to Main mode."
    End If
    If MsgBox(mPrompt & vbCrLf & vbCrLf & "Lab-I will restart to apply the change.", vbQuestion + vbYesNo + vbDefaultButton2, "Change Mode") <> vbYes Then Exit Sub
    SetMiddlewareMode Not MiddlewareMode
    'start a new copy after this one has exited (App.PrevInstance would block an immediate start)
    mExe = App.Path & "\" & App.EXEName & ".exe"
    Shell Environ("ComSpec") & " /c ping -n 4 127.0.0.1 >nul & start """" """ & mExe & """", vbHide
    mSkipExitPrompt = True
    Unload Me
    End
    Exit Sub
errFound:
    MsgBox "Mode saved, but Lab-I could not restart by itself (" & Err.Description & "). Please close and open Lab-I.", vbOKOnly + vbExclamation, "Change Mode"
End Sub

Private Sub mnuControlPanel_Click()
If isConnected Then
    If MsgBox("Do you want to disconnect from machine?", vbQuestion + vbYesNo + vbDefaultButton2, "Disconnect") = vbYes Then
        mnuDisconnect_Click
    End If
End If
    frmControlPanel.Show 1
End Sub

Private Sub mnuDisconnect_Click()
    Unload frmReceiveData
End Sub

Private Sub mnuDoctors_Click()
    HostSA.showphysician
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFlagLists_Click()
    HostSA.showRemarksTemplate
End Sub

Private Sub mnuFormSet_Click()
    HostSA.showformsettings
End Sub

Private Sub mnuHospital_Click()
    HostSA.showHospitalInfo
End Sub

Private Sub mnuInitDTP_Click()
    If MsgBox("DTP LOG " & vbCrLf & ObjDTP.ErrorNum & ": " & ObjDTP.ErrorDesc & vbCrLf & vbCrLf & "Do you want to Initialize Plugin?", vbYesNo + vbQuestion + vbDefaultButton2, "DTP") = vbYes Then
        ObjDTP.ReInit
        If ObjDTP.ErrorNum <> 0 Then
            mnuInitDTP.Visible = True
        Else
            mnuInitDTP.Visible = False
        End If
    End If
End Sub

Private Sub mnuLabExam_Click()
    HostSA.showLabExams
End Sub

Private Sub mnuLabprofiles_Click()
    HostSA.showLabProfiles
End Sub

Private Sub mnuLogbook_Click()
    HostSA.showLogbook
End Sub

Private Sub mnuMedtech_Click()
    HostSA.showmedTech
End Sub

Private Sub mnuNvalues_Click()
    HostSA.showNormalValues
End Sub

Private Sub mnuOtherReports_Click()
    HostSA.showOtherReports
End Sub

Private Sub mnuPathologist_Click()
    HostSA.showPathologist
End Sub

Private Sub mnuPlugTest_Click()
Dim cForms As New ClassForms
Dim rstTemp As ADODB.Recordset
Set rstTemp = myConnection.Execute("select * from labunprocessed")
    If rstTemp.RecordCount > 0 Then
        If cForms.showInputBox("Password", , , , , , True) = "standard" Then frmPlugInTest.Show 1
    Else
        MsgBox "No Record found.", vbOKOnly, App.ProductName
    End If
End Sub

Private Sub mnuprofiles_Click()
    HostSA.showProfileExams
End Sub

Private Sub mnuProgSettings_Click()
Dim EnablePlugin As Integer
Dim PluginName As String
frmSettings.Show 1
EnablePlugin = Val(GetSetting("Lab-I", "Settings", "EnablePlugin", "0"))
If EnablePlugin = 1 Then
    PluginName = Format(GetSetting("Lab-I", "Settings", "PlugIN", ""), ">")
    Set MsgDef = CreateObject(Replace(PluginName, ".DLL", ".MessageDef"))
    InstrumentID = MsgDef.InstrumentID
    MessageFormat = MsgDef.MessageFormat
Else
    InstrumentID = "HMX"
    MessageFormat = "STKS"
End If
End Sub

Private Sub mnuReportFormats_Click()
    HostSA.showReportFormats
    setResultReport
    HostSA.wRPTname = RPTName
End Sub

Private Sub mnuRequest_Click()
    frmRequest.Show 1
End Sub

Private Sub mnuSearch_Click()
If mnuSearch.Checked Then
    mnuSearch.Checked = False
    tbrMain.Buttons("Search").Value = tbrUnpressed
    frmSearch.Hide
    If isResultVisible Then frmSearchResult.Hide
    SaveSetting "HostForms", "Forms", "Viewsearch", "0"
    If isConnected Then frmReceiveData.Height = mdiMain.ScaleHeight
Else
    mnuSearch.Checked = True
    tbrMain.Buttons("Search").Value = tbrPressed
    frmSearch.Show
    frmSearch.dtstart.Value = DateAdd("n", -90, Now)
    frmSearch.dtEnd.Value = Now
    SaveSetting "HostForms", "Forms", "Viewsearch", "1"
End If
End Sub

Private Sub mnuSelectionLists_Click()
    HostSA.showSelectionLists
End Sub

Private Sub mnuSendID_Click()
    myCommStat = cs_SendID
    Load frmReceiveData
    If isConnected Then setButtons 0, 1, 0, 0, 0
End Sub

Private Sub mnuSendReceive_Click()
    myCommStat = cs_SendReceive
    Load frmReceiveData
    If isConnected Then setButtons 0, 1, 0, 0, 0
End Sub


Private Sub mnuStat_Click()
On Error Resume Next
    frmStatistics.Show 1
    If Len(Trim(ReturnValue)) > 0 And mnuSearch.Checked Then
        frmSearch.dtstart = Format(ReturnValue, "mm/dd/yyyy 00:00")
        frmSearch.dtEnd = Format(ReturnValue, "mm/dd/yyyy 23:59")
        frmSearch.txtPatientName = ""
        frmSearch.txtSequenceNum = ""
        frmSearch.chkResultDate.Value = 1
        frmSearch.cmdSearch_Click
    End If
End Sub

Private Sub mnuWorklist_Click()
    frmWorkList.Show 1
End Sub

Private Sub mnuWorklistReport_Click()
    HostSA.showWorklistReport
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Search"
        mnuSearch_Click
    Case "Lab Exams"
        mnuLabExam_Click
    Case "Connect"
        mnuConnect_Click
    Case "Exit"
        mnuExit_Click
    Case "Communication"
        mnuComSet_Click
    Case "Disconnect"
        mnuDisconnect_Click
    Case "database"
        mnuBackUp_Click
    Case "medtechs"
        mnuMedtech_Click
    Case "doctors"
        mnuDoctors_Click
    Case "pathologists"
        mnuPathologist_Click
End Select
End Sub

Private Sub tbrMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
    Case "c1"
        mnuConnect_Click
    Case "c2"
        mnuSendReceive_Click
    Case "c3"
        mnuSendID_Click
End Select
End Sub

Private Sub OpenWSBroadcast()
    wsBroadcast.Protocol = sckUDPProtocol
    wsBroadcast.LocalPort = 1007
    wsBroadcast.RemotePort = 1007
    wsBroadcast.RemoteHost = "255.255.255.255"
    wsBroadcast.SendData wsBroadcast.LocalHostName
End Sub

Private Sub CloseWSBroadcast()
    wsBroadcast.Protocol = sckUDPProtocol
    wsBroadcast.LocalPort = 1007
    wsBroadcast.RemotePort = 1007
    wsBroadcast.RemoteHost = "255.255.255.255"
    wsBroadcast.SendData wsBroadcast.LocalHostName
End Sub

Private Sub tmrDailyBackup_Timer()
Dim timeS As String
Dim tempNow As Date
On Error GoTo errFound
If DBDatabase <> "" And isDBconnected Then
    If BackupSched = "" Then
        timeS = 0
    Else
        timeS = DateDiff("s", Format(Now, "dd-MMM-YY HH:mm:ss"), BackupSched)
        If timeS > 30 Then
            tmrDailyBackup.Interval = 30000
            Exit Sub
        End If
    End If
    If timeS <= 0 Then
        tempNow = Now
        SaveSetting "Lab-I", "Database", "DailyBackup", Format(DateAdd("d", 1, Now), "mm/dd/yyyy 23:59")
        BackupSched = GetSetting("Lab-I", "Database", "DailyBackup", "")
        timeS = DateDiff("s", Format(Now, "dd-MMM-YY HH:mm:ss"), BackupSched)
        frmDailyBackup.Show 1
    End If
    If timeS <= 30 Then
        tmrDailyBackup.Interval = timeS * 1000
    End If
End If
errFound:

End Sub

Private Sub wsBroadcast_DataArrival(ByVal bytesTotal As Long)
    Dim IncomingData As String
    wsBroadcast.GetData IncomingData
    If IncomingData = wsBroadcast.LocalHostName Then
        'frmReceiveData.rtbRcvDisplay.Text = frmReceiveData.rtbRcvDisplay.Text & vbCrLf & IncomingData & " is online"
    Else
        frmReceiveData.rtbRcvDisplay.Text = frmReceiveData.rtbRcvDisplay.Text & vbCrLf & IncomingData & " is online"
    End If
End Sub
'
'Private Sub LocateQuery(myIndex As Integer, myvalue As String)
'Dim rstData As ADODB.Recordset
'Set rstData = myConnection.Execute("Select Distinct Resultdate as [Result Date], PatientName as [Patient Name], Accessionnum as [Sequence #],val(Accessionnum), refnum, patientname, newpatientname,Printed, Status from tbLabInstrumentResults where InstrumentID = 'HMX' and  (PatientName like '%" & myvalue & "%' or NewPatientName like '" & myvalue & "%') order by ResultDate desc")
'If rstData.RecordCount > 0 Then
'    wsServer(myIndex).SendData rstData.RecordCount & " Record(s) Found."
'Else
'    wsServer(myIndex).SendData "No Record Found."
'End If
'End Sub
