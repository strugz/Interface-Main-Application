VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.Ocx"
Begin VB.Form frmSettings2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Communication Settings"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4245
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   4245
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5610
      TabIndex        =   0
      Top             =   1500
      Width           =   4155
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Communication Settings 2"
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
         TabIndex        =   1
         Top             =   150
         Width           =   2505
      End
   End
   Begin VB.Frame fraBody 
      BackColor       =   &H006B4B01&
      Caption         =   "General"
      ForeColor       =   &H0000FFFF&
      Height          =   2535
      Index           =   0
      Left            =   60
      TabIndex        =   8
      Top             =   1470
      Width           =   4125
      Begin VB.ComboBox cboComport 
         Height          =   330
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   270
         Width           =   2175
      End
      Begin VB.ComboBox cboBaudRate 
         Height          =   330
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   630
         Width           =   2175
      End
      Begin VB.ComboBox cboParity 
         Height          =   330
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   990
         Width           =   2175
      End
      Begin VB.ComboBox cboDatabits 
         Height          =   330
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1350
         Width           =   2175
      End
      Begin VB.ComboBox cboStopBits 
         Height          =   330
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1710
         Width           =   2175
      End
      Begin VB.ComboBox cboHandShake 
         Height          =   330
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2070
         Width           =   2175
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Communication Port:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   330
         Width           =   1455
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Baud Rate:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   810
         TabIndex        =   19
         Top             =   690
         Width           =   795
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parity:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   1140
         TabIndex        =   18
         Top             =   1080
         Width           =   450
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Bits:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   900
         TabIndex        =   17
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stop Bits:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   4
         Left            =   900
         TabIndex        =   16
         Top             =   1830
         Width           =   690
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Handshaking:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   5
         Left            =   630
         TabIndex        =   15
         Top             =   2160
         Width           =   975
      End
   End
   Begin VB.Frame fraBody 
      BackColor       =   &H006B4B01&
      Caption         =   "Buffers"
      ForeColor       =   &H0000FFFF&
      Height          =   1635
      Index           =   2
      Left            =   60
      TabIndex        =   21
      Top             =   4050
      Width           =   4125
      Begin VB.CheckBox chkValues 
         Appearance      =   0  'Flat
         BackColor       =   &H006B4B01&
         Caption         =   "EOFEnable"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   1110
         TabIndex        =   27
         Top             =   1350
         Width           =   1755
      End
      Begin VB.TextBox txtInbufferSize 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         TabIndex        =   26
         Top             =   300
         Width           =   705
      End
      Begin VB.TextBox txtRtresHold 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         TabIndex        =   25
         Top             =   630
         Width           =   705
      End
      Begin VB.TextBox txtInputLen 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         TabIndex        =   24
         Top             =   960
         Width           =   705
      End
      Begin VB.TextBox txtOutBufferSize 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3180
         TabIndex        =   23
         Top             =   300
         Width           =   705
      End
      Begin VB.TextBox txtStresHold 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3180
         TabIndex        =   22
         Top             =   630
         Width           =   705
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "InBufferSize:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   32
         Top             =   330
         Width           =   945
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OutBufferSize:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   7
         Left            =   2040
         TabIndex        =   31
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RTreshold:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   8
         Left            =   270
         TabIndex        =   30
         Top             =   690
         Width           =   780
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STreshold:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   9
         Left            =   2340
         TabIndex        =   29
         Top             =   690
         Width           =   780
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "InputLen:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   10
         Left            =   360
         TabIndex        =   28
         Top             =   990
         Width           =   660
      End
   End
   Begin VB.Frame fraBody 
      BackColor       =   &H006B4B01&
      Caption         =   "Hardware"
      ForeColor       =   &H0000FFFF&
      Height          =   1605
      Index           =   3
      Left            =   60
      TabIndex        =   2
      Top             =   5730
      Width           =   4125
      Begin MSCommLib.MSComm mscTest 
         Left            =   3570
         Top             =   1020
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.CheckBox chkValues 
         Appearance      =   0  'Flat
         BackColor       =   &H006B4B01&
         Caption         =   "NullDiscard"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   6
         Top             =   540
         Width           =   1755
      End
      Begin VB.CheckBox chkValues 
         Appearance      =   0  'Flat
         BackColor       =   &H006B4B01&
         Caption         =   "RTSEnable"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   5
         Top             =   840
         Width           =   1755
      End
      Begin VB.CheckBox chkValues 
         Appearance      =   0  'Flat
         BackColor       =   &H006B4B01&
         Caption         =   "DTREnable"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   4
         Top             =   1140
         Width           =   1755
      End
      Begin VB.TextBox txtParityReplace 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1380
         TabIndex        =   3
         Top             =   210
         Width           =   705
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ParityReplace:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   11
         Left            =   300
         TabIndex        =   7
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.Frame fraBody 
      BackColor       =   &H006B4B01&
      Height          =   1365
      Index           =   4
      Left            =   60
      TabIndex        =   36
      Top             =   60
      Width           =   4125
      Begin VB.CheckBox chkFullHandShake 
         Caption         =   "Full Handshake"
         Height          =   225
         Left            =   1710
         TabIndex        =   43
         Top             =   1020
         Width           =   2295
      End
      Begin VB.ComboBox cboControl 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "frmSettings2.frx":0000
         Left            =   1980
         List            =   "frmSettings2.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   240
         Width           =   1905
      End
      Begin VB.CheckBox chkEnabled 
         Appearance      =   0  'Flat
         BackColor       =   &H006B4B01&
         Caption         =   "Enable Port"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   40
         Top             =   990
         Width           =   1275
      End
      Begin VB.CommandButton cmdInstrument 
         Caption         =   "..."
         Height          =   285
         Left            =   3570
         TabIndex        =   39
         Top             =   660
         Width           =   315
      End
      Begin VB.TextBox txtInstrument 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   630
         Width           =   2565
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Communication Control:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   13
         Left            =   150
         TabIndex        =   41
         Top             =   270
         Width           =   1680
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   12
         Left            =   150
         TabIndex        =   37
         Top             =   660
         Width           =   795
      End
   End
   Begin VB.Frame fraBody 
      BackColor       =   &H006B4B01&
      Height          =   645
      Index           =   1
      Left            =   60
      TabIndex        =   33
      Top             =   7380
      Width           =   4125
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   315
         Left            =   2730
         TabIndex        =   35
         Top             =   210
         Width           =   1245
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   315
         Left            =   1410
         TabIndex        =   34
         Top             =   210
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmSettings2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mySettings As New ComSettings

Private Sub cboControl_Change()
AssignValues
End Sub

Private Sub cboControl_Click()
AssignValues
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub fillcbo()
Dim i As Integer
With cboComport
Select Case getVersion
    Case "WIN32", "Failed"
        For i = 1 To 16
            .AddItem "COM" & i
        Next
    Case "NT"
        For i = 1 To 16
            If COMAvailable2(i) Then
                .AddItem "COM" & i
            End If
        Next
End Select
End With
With cboBaudRate
    .AddItem "110"
    .AddItem "300"
    .AddItem "1200"
    .AddItem "2400"
    .AddItem "4800"
    .AddItem "9600"
    .AddItem "19200"
    .AddItem "38400"
    .AddItem "57600"
    .AddItem "115200"
    .AddItem "230400"
    .AddItem "460800"
    .AddItem "921600"
End With
With cboParity
    .AddItem "ODD"
    .AddItem "EVEN"
    .AddItem "NONE"
End With
With cboDatabits
    .AddItem "5"
    .AddItem "6"
    .AddItem "7"
    .AddItem "8"
End With
With cboStopBits
    .AddItem "1"
    .AddItem "1.5"
    .AddItem "2"
End With
With cboHandShake
    .AddItem "0-ComNone"
    .AddItem "1-ComXonXoff"
    .AddItem "2-ComRTS"
End With
With cboControl
    .AddItem "Control 1"
    .AddItem "Control 2"
    .AddItem "Control 3"
    .AddItem "Control 4"
    .ListIndex = 0
End With
End Sub

Private Sub cmdInstrument_Click()
frmSelectPlugins.Show 1
txtInstrument.Text = returnValue
End Sub

Private Sub cmdSave_Click()
    If Trim(cboComport.Text) = "" Then
        MsgBox "Unable to save communication settings due to port cannot be detected.", vbOKOnly + vbDefaultButton2 + vbInformation, "Communication Settings"
    Else
        SaveValues
        MsgBox "Values has been successfully saved!", vbOKOnly, "Communication Settings"
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
setFormColors Me
    fillcbo
    AssignValues
End Sub

Private Function findCboIndex(cboTemp As ComboBox, texttofind As String) As Integer
Dim ctr As Integer
Dim tempIndex As Integer
tempIndex = -1
With cboTemp
    For ctr = 0 To .ListCount - 1
        .ListIndex = ctr
        If .Text = texttofind Then
            tempIndex = .ListIndex
            Exit For
        End If
    Next ctr
End With
findCboIndex = tempIndex
End Function

Private Sub AssignValues()
    If cboControl.Text = "Control 1" Then
        cmdInstrument.Enabled = False
        txtInstrument.Enabled = False
        If Val(GetSetting("Lab-I", "Settings", "EnablePlugin", "0")) = 1 Then
            txtInstrument.Text = Format(GetSetting("Lab-I", "Settings", "PlugIN", ""), ">")
        Else
            txtInstrument.Text = "STKS/HMX/LH"
        End If
    Else
        cmdInstrument.Enabled = True
        txtInstrument.Enabled = True
        If Len(Trim(GetSetting("ComSet", cboControl.Text, "Plug-Ins", ""))) = 0 Then
            txtInstrument.Text = "STKS/HMX/LH"
        Else
            txtInstrument.Text = GetSetting("ComSet", cboControl.Text, "Plug-Ins", "")
        End If
    End If
    cboComport.ListIndex = findCboIndex(cboComport, GetSetting("ComSet", cboControl.Text, "ComPort", "COM1"))
    cboBaudRate.ListIndex = findCboIndex(cboBaudRate, GetSetting("ComSet", cboControl.Text, "BaudRate", "9600"))
    cboParity.ListIndex = findCboIndex(cboParity, GetSetting("ComSet", cboControl.Text, "Parity", "NONE"))
    cboDatabits.ListIndex = findCboIndex(cboDatabits, GetSetting("ComSet", cboControl.Text, "DataBits", "8"))
    cboStopBits.ListIndex = findCboIndex(cboStopBits, GetSetting("ComSet", cboControl.Text, "StopBits", "1"))
    cboHandShake.ListIndex = findCboIndex(cboHandShake, GetSetting("ComSet", cboControl.Text, "HandShake", "0-ComNone"))
    txtInbufferSize.Text = GetSetting("ComSet", cboControl.Text, "InbufferSize", "1024")
    txtOutBufferSize.Text = GetSetting("ComSet", cboControl.Text, "OutbufferSize", "512")
    txtRtresHold.Text = GetSetting("ComSet", cboControl.Text, "RTresHold", "1")
    txtStresHold.Text = GetSetting("ComSet", cboControl.Text, "STresHold", "1")
    txtInputLen.Text = GetSetting("ComSet", cboControl.Text, "InputLen", "1")
    txtParityReplace.Text = GetSetting("ComSet", cboControl.Text, "ParityReplace", "?")
    chkValues(0).Value = Val(GetSetting("ComSet", cboControl.Text, "EOFEnable", "0"))
    chkValues(1).Value = Val(GetSetting("ComSet", cboControl.Text, "NullDiscard", "1"))
    chkValues(2).Value = Val(GetSetting("ComSet", cboControl.Text, "RTSEnable", "1"))
    chkValues(3).Value = Val(GetSetting("ComSet", cboControl.Text, "DTREnable", "1"))
    chkEnabled.Value = Val(GetSetting("ComSet", cboControl.Text, "Enabled", "0"))
    chkFullHandShake.Value = Val(GetSetting("ComSet", cboControl.Text, "FullHandshake", "0"))
End Sub

Private Sub SaveValues()
    SaveSetting "ComSet", cboControl.Text, "ComPort", cboComport.Text
    SaveSetting "ComSet", cboControl.Text, "BaudRate", cboBaudRate.Text
    SaveSetting "ComSet", cboControl.Text, "Parity", cboParity.Text
    SaveSetting "ComSet", cboControl.Text, "DataBits", cboDatabits.Text
    SaveSetting "ComSet", cboControl.Text, "StopBits", cboStopBits.Text
    SaveSetting "ComSet", cboControl.Text, "HandShake", cboHandShake.Text
    SaveSetting "ComSet", cboControl.Text, "InbufferSize", txtInbufferSize.Text
    SaveSetting "ComSet", cboControl.Text, "OutbufferSize", txtOutBufferSize.Text
    SaveSetting "ComSet", cboControl.Text, "RTresHold", txtRtresHold.Text
    SaveSetting "ComSet", cboControl.Text, "STresHold", txtStresHold.Text
    SaveSetting "ComSet", cboControl.Text, "InputLen", txtInputLen.Text
    SaveSetting "ComSet", cboControl.Text, "ParityReplace", txtParityReplace
    SaveSetting "ComSet", cboControl.Text, "EOFEnable", chkValues(0).Value
    SaveSetting "ComSet", cboControl.Text, "NullDiscard", chkValues(1).Value
    SaveSetting "ComSet", cboControl.Text, "RTSEnable", chkValues(2).Value
    SaveSetting "ComSet", cboControl.Text, "DTREnable", chkValues(3).Value
    SaveSetting "ComSet", cboControl.Text, "Enabled", chkEnabled.Value
    SaveSetting "ComSet", cboControl.Text, "Plug-Ins", txtInstrument.Text
    SaveSetting "ComSet", cboControl.Text, "FullHandshake", chkFullHandShake.Value
End Sub

Private Sub lblText_DblClick(Index As Integer)
Dim i As Integer
With cboComport
    .Clear
    For i = 1 To 4
        .AddItem "COM" & i
    Next
End With
End Sub

Private Sub txtInstrument_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    txtInstrument.Text = "STKS\HMX\LH"
End If
End Sub

Private Function COMAvailable2(COMNum As Integer) As Boolean
On Error GoTo errFound
With mscTest
    .CommPort = COMNum
    If .PortOpen Then
        COMAvailable2 = False
    Else
        COMAvailable2 = True
        .PortOpen = True
        .PortOpen = False
    End If
End With
errFound:
    If Err.Number <> 0 Then
        COMAvailable2 = False
    End If
End Function

