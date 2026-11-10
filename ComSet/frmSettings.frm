VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   0  'None
   Caption         =   "Communication Settings"
   ClientHeight    =   6840
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   34
      Top             =   -60
      Width           =   4125
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Communication Settings"
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
         TabIndex        =   35
         Top             =   150
         Width           =   2325
      End
   End
   Begin VB.Frame fraBody 
      BackColor       =   &H006B4B01&
      Caption         =   "Hardware"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1605
      Index           =   3
      Left            =   0
      TabIndex        =   17
      Top             =   4650
      Width           =   4125
      Begin VB.TextBox txtParityReplace 
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
         Height          =   315
         Left            =   1380
         TabIndex        =   28
         Top             =   210
         Width           =   705
      End
      Begin VB.CheckBox chkValues 
         Appearance      =   0  'Flat
         BackColor       =   &H006B4B01&
         Caption         =   "DTREnable"
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
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   21
         Top             =   1140
         Width           =   1755
      End
      Begin VB.CheckBox chkValues 
         Appearance      =   0  'Flat
         BackColor       =   &H006B4B01&
         Caption         =   "RTSEnable"
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
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   20
         Top             =   840
         Width           =   1755
      End
      Begin VB.CheckBox chkValues 
         Appearance      =   0  'Flat
         BackColor       =   &H006B4B01&
         Caption         =   "NullDiscard"
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
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   19
         Top             =   540
         Width           =   1755
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ParityReplace:"
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
         Index           =   11
         Left            =   300
         TabIndex        =   18
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.Frame fraBody 
      BackColor       =   &H006B4B01&
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2535
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   450
      Width           =   4125
      Begin VB.ComboBox cboHandShake 
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   2070
         Width           =   2175
      End
      Begin VB.ComboBox cboStopBits 
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1710
         Width           =   2175
      End
      Begin VB.ComboBox cboDatabits 
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1350
         Width           =   2175
      End
      Begin VB.ComboBox cboParity 
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   990
         Width           =   2175
      End
      Begin VB.ComboBox cboBaudRate 
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   630
         Width           =   2175
      End
      Begin VB.ComboBox cboComport 
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   270
         Width           =   2175
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Handshaking:"
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
         Index           =   5
         Left            =   630
         TabIndex        =   10
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stop Bits:"
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
         Left            =   900
         TabIndex        =   8
         Top             =   1830
         Width           =   690
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Bits:"
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
         Left            =   900
         TabIndex        =   7
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parity:"
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
         Left            =   1140
         TabIndex        =   6
         Top             =   1080
         Width           =   450
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Baud Rate:"
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
         Left            =   810
         TabIndex        =   5
         Top             =   690
         Width           =   795
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Communication Port:"
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
         Left            =   180
         TabIndex        =   4
         Top             =   330
         Width           =   1455
      End
   End
   Begin VB.Frame fraBody 
      BackColor       =   &H006B4B01&
      Caption         =   "Buffers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1635
      Index           =   2
      Left            =   0
      TabIndex        =   9
      Top             =   3000
      Width           =   4125
      Begin VB.TextBox txtStresHold 
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
         Height          =   315
         Left            =   3180
         TabIndex        =   27
         Top             =   630
         Width           =   705
      End
      Begin VB.TextBox txtOutBufferSize 
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
         Height          =   315
         Left            =   3180
         TabIndex        =   26
         Top             =   300
         Width           =   705
      End
      Begin VB.TextBox txtInputLen 
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
         Height          =   315
         Left            =   1110
         TabIndex        =   25
         Top             =   960
         Width           =   705
      End
      Begin VB.TextBox txtRtresHold 
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
         Height          =   315
         Left            =   1110
         TabIndex        =   24
         Top             =   630
         Width           =   705
      End
      Begin VB.TextBox txtInbufferSize 
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
         Height          =   315
         Left            =   1110
         TabIndex        =   23
         Top             =   300
         Width           =   705
      End
      Begin VB.CheckBox chkValues 
         Appearance      =   0  'Flat
         BackColor       =   &H006B4B01&
         Caption         =   "EOFEnable"
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
         Height          =   195
         Index           =   0
         Left            =   1110
         TabIndex        =   16
         Top             =   1350
         Width           =   1755
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "InputLen:"
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
         Index           =   10
         Left            =   360
         TabIndex        =   15
         Top             =   990
         Width           =   660
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STreshold:"
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
         Index           =   9
         Left            =   2340
         TabIndex        =   14
         Top             =   690
         Width           =   780
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RTreshold:"
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
         Index           =   8
         Left            =   270
         TabIndex        =   13
         Top             =   690
         Width           =   780
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OutBufferSize:"
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
         Index           =   7
         Left            =   2040
         TabIndex        =   12
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "InBufferSize:"
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
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   330
         Width           =   945
      End
   End
   Begin VB.Frame fraBody 
      BackColor       =   &H006B4B01&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   6180
      Width           =   4125
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
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
         Left            =   1410
         TabIndex        =   3
         Top             =   210
         Width           =   1245
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
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
         Left            =   2730
         TabIndex        =   2
         Top             =   210
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mySettings As New ComSettings

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub fillcbo()
Dim i As Integer
With cboComport
Select Case getVersion
    Case "WIN32", "Failed"
        For i = 1 To 4
            .AddItem "COM" & i
        Next
    Case "NT"
        For i = 1 To 16
            If COMAvailable(i) Then
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
End Sub

Private Sub cmdSave_Click()
    SaveValues
    MsgBox "Values has been successfully saved!", vbOKOnly, "Communication Settings"
End Sub

Private Sub Form_Load()
setFormColors Me
    fillcbo
    AssignValues
End Sub

Private Function findCboIndex(cboTemp As ComboBox, texttofind As String) As Integer
Dim ctr As Integer
With cboTemp
    For ctr = 0 To .ListCount - 1
        .ListIndex = ctr
        If .Text = texttofind Then
            findCboIndex = .ListIndex
            Exit Function
        End If
    Next ctr
End With
End Function

Private Sub AssignValues()
    cboComport.ListIndex = findCboIndex(cboComport, GetSetting("ComSet", "PortSettings", "ComPort", "COM1"))
    cboBaudRate.ListIndex = findCboIndex(cboBaudRate, GetSetting("ComSet", "PortSettings", "BaudRate", "9600"))
    cboParity.ListIndex = findCboIndex(cboParity, GetSetting("ComSet", "PortSettings", "Parity", "NONE"))
    cboDatabits.ListIndex = findCboIndex(cboDatabits, GetSetting("ComSet", "PortSettings", "DataBits", "8"))
    cboStopBits.ListIndex = findCboIndex(cboStopBits, GetSetting("ComSet", "PortSettings", "StopBits", "1"))
    cboHandShake.ListIndex = findCboIndex(cboHandShake, GetSetting("ComSet", "PortSettings", "HandShake", "0-ComNone"))
    txtInbufferSize.Text = GetSetting("ComSet", "PortSettings", "InbufferSize", "1024")
    txtOutBufferSize.Text = GetSetting("ComSet", "PortSettings", "OutbufferSize", "512")
    txtRtresHold.Text = GetSetting("ComSet", "PortSettings", "RTresHold", "1")
    txtStresHold.Text = GetSetting("ComSet", "PortSettings", "STresHold", "1")
    txtInputLen.Text = GetSetting("ComSet", "PortSettings", "InputLen", "1")
    txtParityReplace.Text = GetSetting("ComSet", "PortSettings", "ParityReplace", "?")
    chkValues(0).Value = Val(GetSetting("ComSet", "PortSettings", "EOFEnable", "0"))
    chkValues(1).Value = Val(GetSetting("ComSet", "PortSettings", "NullDiscard", "1"))
    chkValues(2).Value = Val(GetSetting("ComSet", "PortSettings", "RTSEnable", "1"))
    chkValues(3).Value = Val(GetSetting("ComSet", "PortSettings", "DTREnable", "1"))
End Sub

Private Sub SaveValues()
    SaveSetting "ComSet", "PortSettings", "ComPort", cboComport.Text
    SaveSetting "ComSet", "PortSettings", "BaudRate", cboBaudRate.Text
    SaveSetting "ComSet", "PortSettings", "Parity", cboParity.Text
    SaveSetting "ComSet", "PortSettings", "DataBits", cboDatabits.Text
    SaveSetting "ComSet", "PortSettings", "StopBits", cboStopBits.Text
    SaveSetting "ComSet", "PortSettings", "HandShake", cboHandShake.Text
    SaveSetting "ComSet", "PortSettings", "InbufferSize", txtInbufferSize.Text
    SaveSetting "ComSet", "PortSettings", "OutbufferSize", txtOutBufferSize.Text
    SaveSetting "ComSet", "PortSettings", "RTresHold", txtRtresHold.Text
    SaveSetting "ComSet", "PortSettings", "STresHold", txtStresHold.Text
    SaveSetting "ComSet", "PortSettings", "InputLen", txtInputLen.Text
    SaveSetting "ComSet", "PortSettings", "ParityReplace", txtParityReplace
    SaveSetting "ComSet", "PortSettings", "EOFEnable", chkValues(0).Value
    SaveSetting "ComSet", "PortSettings", "NullDiscard", chkValues(1).Value
    SaveSetting "ComSet", "PortSettings", "RTSEnable", chkValues(2).Value
    SaveSetting "ComSet", "PortSettings", "DTREnable", chkValues(3).Value
End Sub

