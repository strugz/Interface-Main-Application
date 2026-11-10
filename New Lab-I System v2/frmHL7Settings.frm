VERSION 5.00
Begin VB.Form frmHL7Settings 
   Caption         =   "HL7 Settings"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   4110
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAutoWrite 
      Caption         =   "Auto Write HL7 Message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1080
      TabIndex        =   29
      Top             =   4920
      Width           =   2835
   End
   Begin VB.CheckBox chkAutoSend 
      Caption         =   "Auto Send Request to Instrument"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1110
      TabIndex        =   28
      Top             =   3270
      Width           =   2835
   End
   Begin VB.TextBox txtExtOut 
      Height          =   285
      Left            =   1110
      TabIndex        =   26
      Top             =   4590
      Width           =   2865
   End
   Begin VB.TextBox txtDirOut 
      Height          =   285
      Left            =   1110
      TabIndex        =   24
      Top             =   4260
      Width           =   2865
   End
   Begin VB.TextBox txtFileExt 
      Height          =   285
      Left            =   1110
      TabIndex        =   22
      Top             =   2910
      Width           =   2865
   End
   Begin VB.TextBox txtFileDir 
      Height          =   285
      Left            =   1110
      TabIndex        =   21
      Top             =   2580
      Width           =   2865
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5220
      TabIndex        =   17
      Top             =   840
      Width           =   2235
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5220
      TabIndex        =   16
      Top             =   510
      Width           =   2235
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5220
      TabIndex        =   15
      Top             =   180
      Width           =   2235
   End
   Begin VB.TextBox txtEC 
      Height          =   285
      Left            =   1710
      TabIndex        =   11
      Top             =   510
      Width           =   2235
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   435
      Left            =   1230
      TabIndex        =   10
      Top             =   5460
      Width           =   1305
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   435
      Left            =   2640
      TabIndex        =   9
      Top             =   5460
      Width           =   1305
   End
   Begin VB.TextBox txtVersion 
      Height          =   285
      Left            =   1710
      TabIndex        =   8
      Top             =   1590
      Width           =   2235
   End
   Begin VB.TextBox txtRA 
      Height          =   285
      Left            =   1710
      TabIndex        =   7
      Top             =   1230
      Width           =   2235
   End
   Begin VB.TextBox txtSA 
      Height          =   285
      Left            =   1710
      TabIndex        =   6
      Top             =   870
      Width           =   2235
   End
   Begin VB.TextBox txtFS 
      Height          =   285
      Left            =   1710
      TabIndex        =   5
      Top             =   150
      Width           =   2235
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      Caption         =   "File Extension:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   11
      Left            =   60
      TabIndex        =   27
      Top             =   4620
      Width           =   1035
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      Caption         =   "File Directory:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   10
      Left            =   120
      TabIndex        =   25
      Top             =   4320
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "File Writer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   150
      TabIndex        =   23
      Top             =   3960
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File Reader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      Caption         =   "File Extension:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   60
      TabIndex        =   19
      Top             =   2940
      Width           =   1035
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      Caption         =   "File Directory:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   990
   End
   Begin VB.Line Line1 
      X1              =   150
      X2              =   3960
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      Caption         =   "Local Port:"
      Height          =   195
      Index           =   7
      Left            =   4380
      TabIndex        =   14
      Top             =   870
      Width           =   765
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      Caption         =   "Remote Port:"
      Height          =   195
      Index           =   6
      Left            =   4200
      TabIndex        =   13
      Top             =   510
      Width           =   930
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      Caption         =   "Remote Host:"
      Height          =   195
      Index           =   5
      Left            =   4170
      TabIndex        =   12
      Top             =   210
      Width           =   975
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      Caption         =   "HL7 Version:"
      Height          =   195
      Index           =   4
      Left            =   690
      TabIndex        =   4
      Top             =   1620
      Width           =   915
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      Caption         =   "Receiving Application:"
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   3
      Top             =   1260
      Width           =   1590
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      Caption         =   "Sending Application:"
      Height          =   195
      Index           =   2
      Left            =   210
      TabIndex        =   2
      Top             =   900
      Width           =   1455
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      Caption         =   "Encoding Characters:"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   540
      Width           =   1530
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      Caption         =   "Field Separator:"
      Height          =   195
      Index           =   0
      Left            =   570
      TabIndex        =   0
      Top             =   180
      Width           =   1110
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0FF&
      BorderWidth     =   2
      X1              =   180
      X2              =   3990
      Y1              =   2100
      Y2              =   2100
   End
End
Attribute VB_Name = "frmHL7Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    SaveSetting "LAB-i", "HL7 Settings", "Field Separator", txtFS.Text
    SaveSetting "LAB-i", "HL7 Settings", "Encoding Characters", txtEC.Text
    SaveSetting "LAB-i", "HL7 Settings", "Sending Application", txtSA.Text
    SaveSetting "LAB-i", "HL7 Settings", "Receiving Application", txtRA.Text
    SaveSetting "LAB-i", "HL7 Settings", "HL7 Version", txtVersion.Text
    SaveSetting "LAB-i", "HL7 Settings", "FRDIR", txtFileDir.Text
    SaveSetting "LAB-i", "HL7 Settings", "FREXT", txtFileExt.Text
    SaveSetting "LAB-i", "HL7 Settings", "FWDIR", txtDirOut.Text
    SaveSetting "LAB-i", "HL7 Settings", "FWEXT", txtExtOut.Text
    SaveSetting "LAB-i", "HL7 Settings", "AutoSendRequest", chkAutoSend.Value
    SaveSetting "LAB-i", "HL7 Settings", "AutoWriteResult", chkAutoWrite.Value
End Sub

Private Sub Form_Load()
    txtFS.Text = GetSetting("LAB-i", "HL7 Settings", "Field Separator", "|")
    txtEC.Text = GetSetting("LAB-i", "HL7 Settings", "Encoding Characters", "^~\&")
    txtSA.Text = GetSetting("LAB-i", "HL7 Settings", "Sending Application", "LAB-Iv" & App.Major)
    txtRA.Text = GetSetting("LAB-i", "HL7 Settings", "Receiving Application", "MDMPI")
    txtVersion.Text = GetSetting("LAB-i", "HL7 Settings", "HL7 Version", "2.5.1")
    txtFileDir.Text = GetSetting("LAB-i", "HL7 Settings", "FRDIR", App.Path & "\Received")
    txtFileExt.Text = GetSetting("LAB-i", "HL7 Settings", "FREXT", "txt")
    txtDirOut.Text = GetSetting("LAB-i", "HL7 Settings", "FWDIR", App.Path & "\HostSpooler")
    txtExtOut.Text = GetSetting("LAB-i", "HL7 Settings", "FWEXT", "txt")
    chkAutoSend.Value = Val(GetSetting("LAB-i", "HL7 Settings", "AutoSendRequest", 0))
    chkAutoWrite.Value = Val(GetSetting("LAB-i", "HL7 Settings", "AutoWriteResult", 0))
End Sub
