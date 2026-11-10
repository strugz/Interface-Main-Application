VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmPIDsearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PID Search"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7620
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7620
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   510
      Top             =   -30
   End
   Begin MSWinsockLib.Winsock wsTest 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   8580
      TabIndex        =   0
      Top             =   1800
      Width           =   1545
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PID Search"
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
         Width           =   1065
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   2805
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   7485
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   2565
         Left            =   90
         TabIndex        =   4
         Top             =   180
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   4524
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   12615680
         ForeColorFixed  =   16777215
         BackColorBkg    =   12632256
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
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
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblText 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No Record Found..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Top             =   1410
         Width           =   7395
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   2265
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Top             =   2910
      Width           =   7485
      Begin VB.TextBox txtResultdate 
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
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1830
         Width           =   2565
      End
      Begin VB.TextBox txtRequestDate 
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
         Left            =   4770
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   180
         Width           =   2565
      End
      Begin VB.TextBox txtRoom 
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
         Left            =   6120
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   30
         Top             =   1830
         Width           =   1215
      End
      Begin VB.TextBox txtWard 
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
         Left            =   4770
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   29
         Top             =   1830
         Width           =   915
      End
      Begin VB.TextBox txtPhysician 
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
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   28
         Top             =   1500
         Width           =   2565
      End
      Begin VB.TextBox txtLastName 
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
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   17
         Top             =   510
         Width           =   2565
      End
      Begin VB.TextBox txtFirstName 
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
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   16
         Top             =   840
         Width           =   2565
      End
      Begin VB.TextBox txtMiddleName 
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
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   15
         Top             =   1170
         Width           =   2565
      End
      Begin VB.TextBox txtAge 
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
         Left            =   4770
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   14
         Top             =   1170
         Width           =   975
      End
      Begin VB.TextBox txtSex 
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
         Left            =   6120
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   13
         Top             =   1170
         Width           =   1215
      End
      Begin VB.TextBox txtPatientID 
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
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   12
         Top             =   180
         Width           =   2565
      End
      Begin VB.TextBox txtCivilStatus 
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
         Left            =   4770
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   11
         Top             =   840
         Width           =   2565
      End
      Begin VB.TextBox txtBirthDate 
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
         Left            =   4770
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   510
         Width           =   2565
      End
      Begin VB.TextBox txtSpecimen 
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
         Left            =   4770
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1500
         Width           =   2565
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Left            =   90
         TabIndex        =   35
         Top             =   1830
         Width           =   1020
      End
      Begin VB.Label lblRequestDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Request Date:"
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
         Left            =   3750
         TabIndex        =   34
         Top             =   210
         Width           =   1020
      End
      Begin VB.Label lblRoom 
         AutoSize        =   -1  'True
         BackColor       =   &H00FCE2C7&
         BackStyle       =   0  'Transparent
         Caption         =   "Room:"
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
         Left            =   5670
         TabIndex        =   32
         Top             =   1860
         Width           =   450
      End
      Begin VB.Label lblward 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ward:"
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
         Left            =   3750
         TabIndex        =   31
         Top             =   1860
         Width           =   975
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Physician:"
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
         Left            =   -120
         TabIndex        =   27
         Top             =   1530
         Width           =   1215
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
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
         Left            =   60
         TabIndex        =   26
         Top             =   510
         Width           =   990
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FCE2C7&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
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
         Left            =   30
         TabIndex        =   25
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name:"
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
         Left            =   -150
         TabIndex        =   24
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblAge 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Age:"
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
         Left            =   3750
         TabIndex        =   23
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackColor       =   &H00FCE2C7&
         BackStyle       =   0  'Transparent
         Caption         =   "Sex:"
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
         Left            =   5760
         TabIndex        =   22
         Top             =   1200
         Width           =   330
      End
      Begin VB.Label lblPatientID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Patient ID:"
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
         Left            =   60
         TabIndex        =   21
         Top             =   210
         Width           =   975
      End
      Begin VB.Label lblCivilStatus 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FCE2C7&
         BackStyle       =   0  'Transparent
         Caption         =   "Civil Status:"
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
         Left            =   3720
         TabIndex        =   20
         Top             =   900
         Width           =   1020
      End
      Begin VB.Label lblBirthDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date:"
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
         Left            =   3750
         TabIndex        =   19
         Top             =   540
         Width           =   990
      End
      Begin VB.Label lblSpecimen 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FCE2C7&
         BackStyle       =   0  'Transparent
         Caption         =   "Specimen:"
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
         Left            =   3510
         TabIndex        =   18
         Top             =   1500
         Width           =   1230
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   705
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   5220
      Width           =   7485
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   8
         Top             =   210
         Width           =   1395
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5970
         TabIndex        =   7
         Top             =   210
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmPIDsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstPatient As ADODB.Recordset
Private ySearchID As String
Private yReturnValue As String
Private yExDataBase As Boolean
Private blnConnected As Boolean
Private RequestDone As Boolean
Private DataReceived As String
Private TimeOutCtr As Integer
Private mTimeOut As Integer

Public Property Let wSearchString(ByVal vNewValue As String)
    ySearchID = vNewValue
End Property

Public Property Get wReturnValue() As String
    wReturnValue = yReturnValue
End Property

Public Property Let wExDataBase(ByVal vNewValue As Boolean)
    yExDataBase = vNewValue
End Property

Private Sub cmdClose_Click()
    yReturnValue = ""
    Set rstPatient = Nothing
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    yReturnValue = Trim(txtPatientID.Text) & "|" & Trim(txtLastName.Text) & "|" & _
    Trim(txtFirstName.Text) & "|" & _
    Trim(txtMiddleName.Text) & "|" & _
    Trim(txtBirthDate.Text) & "|" & _
    Trim(txtCivilStatus.Text) & "|" & _
    Trim(txtAge.Text) & "|" & _
    Trim(txtSex.Text) & "|" & _
    Trim(txtSpecimen.Text) & "|" & _
    Trim(txtPhysician.Text) & "|" & _
    Trim(txtWard.Text) & "|" & _
    Trim(txtRoom.Text) & "|" & _
    Trim(txtRequestDate.Text) & "|" & _
    Trim(txtResultdate.Text)
    Unload Me
End Sub

Private Sub Form_Load()
Dim tempStr As String
setFormColors Me
RequestDone = True
yReturnValue = ""
lblPatientID.ForeColor = lblText(0).ForeColor
lblAge.ForeColor = lblText(0).ForeColor
lblSex.ForeColor = lblText(0).ForeColor
lblBirthDate.ForeColor = lblText(0).ForeColor
lblCivilStatus.ForeColor = lblText(0).ForeColor
lblSpecimen.ForeColor = lblText(0).ForeColor
lblWard.ForeColor = lblText(0).ForeColor
lblRoom.ForeColor = lblText(0).ForeColor
lblRequestDate.ForeColor = lblText(0).ForeColor
    If yExDataBase Then
        mTimeOut = Val(GetSetting("LAB-I", "OBISettings", "QueryTimeOut", "5"))
        If GetSetting("LAB-I", "OBISettings", "ServerType", "") = "MS SQL Server 2000" Then
            Set rstPatient = myCon.Execute("select distinct " & GetSetting("LAB-I", "OBISettings", "FLD_1", "''") & " as [Patient ID],rtrim(" & GetSetting("LAB-I", "OBISettings", "FLD_3", "''") & ") + ', ' + rtrim(" & GetSetting("LAB-I", "OBISettings", "FLD_4", "''") & ") + ' ' + rtrim(" & GetSetting("LAB-I", "OBISettings", "FLD_5", "''") & ") as [Patient Name] " & _
            "," & GetSetting("LAB-I", "OBISettings", "FLD_17", "''") & ", " & GetSetting("LAB-I", "OBISettings", "FLD_18", "''") & "," & GetSetting("LAB-I", "OBISettings", "FLD_19", "''") & "," & GetSetting("LAB-I", "OBISettings", "FLD_6", "''") & "," & GetSetting("LAB-I", "OBISettings", "FLD_7", "''") & "," & GetSetting("LAB-I", "OBISettings", "FLD_3", "''") & _
            "," & GetSetting("LAB-I", "OBISettings", "FLD_4", "''") & "," & GetSetting("LAB-I", "OBISettings", "FLD_5", "''") & ",''," & GetSetting("LAB-I", "OBISettings", "FLD_8", "''") & "," & GetSetting("LAB-I", "OBISettings", "FLD_9", "''") & _
            "," & GetSetting("LAB-I", "OBISettings", "FLD_2", "''") & "," & GetSetting("LAB-I", "OBISettings", "FLD_10", "''") & ",''," & GetSetting("LAB-I", "OBISettings", "FLD_11", "''") & "," & GetSetting("LAB-I", "OBISettings", "FLD_12", "''") & "," & GetSetting("LAB-I", "OBISettings", "FLD_13", "''") & _
            "," & GetSetting("LAB-I", "OBISettings", "FLD_14", "''") & " " & _
            " from " & GetSetting("LAB-I", "OBISettings", "Table", "") & " where ltrim(rtrim(" & GetSetting("LAB-I", "OBISettings", "FLD_1", "''") & ")) = '" & _
            Trim(ySearchID) & "' and " & GetSetting("LAB-I", "OBISettings", "FLD_1", "''") & " <> ''" & " " & GetSetting("LAB-I", "OBISettings", "Criteria", ""))
        ElseIf GetSetting("LAB-I", "OBISettings", "ServerType", "") = "MySQL" Then
            Set rstPatient = myCon.Execute("select distinct " & GetSetting("LAB-I", "OBISettings", "FLD_1", "''") & " as `Patient ID`,trim(" & GetSetting("LAB-I", "OBISettings", "FLD_3", "''") & ") + ', ' + trim(" & GetSetting("LAB-I", "OBISettings", "FLD_4", "''") & ") + ' ' + trim(" & GetSetting("LAB-I", "OBISettings", "FLD_5", "''") & ") as `Patient Name` " & _
            "," & GetSetting("LAB-I", "OBISettings", "FLD_17", "''") & ", " & GetSetting("LAB-I", "OBISettings", "FLD_18", "''") & "," & GetSetting("LAB-I", "OBISettings", "FLD_19", "''") & "," & GetSetting("LAB-I", "OBISettings", "FLD_6", "''") & "," & GetSetting("LAB-I", "OBISettings", "FLD_7", "''") & "," & GetSetting("LAB-I", "OBISettings", "FLD_3", "''") & _
            "," & GetSetting("LAB-I", "OBISettings", "FLD_4", "''") & "," & GetSetting("LAB-I", "OBISettings", "FLD_5", "''") & ",''," & GetSetting("LAB-I", "OBISettings", "FLD_8", "''") & "," & GetSetting("LAB-I", "OBISettings", "FLD_9", "''") & _
            "," & GetSetting("LAB-I", "OBISettings", "FLD_2", "''") & "," & GetSetting("LAB-I", "OBISettings", "FLD_10", "''") & ",''," & GetSetting("LAB-I", "OBISettings", "FLD_11", "''") & "," & GetSetting("LAB-I", "OBISettings", "FLD_12", "''") & "," & GetSetting("LAB-I", "OBISettings", "FLD_13", "''") & _
            "," & GetSetting("LAB-I", "OBISettings", "FLD_14", "''") & " " & _
            " from " & GetSetting("LAB-I", "OBISettings", "Table", "") & " where ltrim(trim(" & GetSetting("LAB-I", "OBISettings", "FLD_1", "''") & ")) = '" & _
            Trim(ySearchID) & "' and " & GetSetting("LAB-I", "OBISettings", "FLD_1", "''") & " <> ''" & " " & GetSetting("LAB-I", "OBISettings", "Criteria", ""))
        ElseIf GetSetting("LAB-I", "OBISettings", "ServerType", "") = "REST API" Then
            GenerateTestPacket
            Do While Not RequestDone
                DoEvents
            Loop
            If GetError(DataReceived, 1) = "0" Then
                ClearValues
                assignValues 0
                cmdSelect_Click
            Else
                ClearValues
                yReturnValue = ""
                Unload Me
            End If
        End If
    Else
        Set rstPatient = myConnection.Execute("select distinct patientID as [Patient ID],Lastname + ', ' + firstname + ' ' + middlename as [Patient Name] " & _
        ",BirthDate, CivilStatus,Specimen,Age,Sex,Lastname,Firstname,Middlename,RefNum,Ward,Room " & _
        " from tbpatientmaster where patientid = '" & _
        ySearchID & "' and patientid <> ''")
    End If
    If GetSetting("LAB-I", "OBISettings", "ServerType", "") = "REST API" Then
    Else
        If rstPatient.RecordCount > 1 Then
            Set msfgResult.DataSource = rstPatient
            setFlexGrid
            ClearValues
            assignValues 1
        ElseIf rstPatient.RecordCount = 1 Then
            Set msfgResult.DataSource = rstPatient
            setFlexGrid
            ClearValues
            assignValues 1
            cmdSelect_Click
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub setFlexGrid()
With msfgResult
    .ColWidth(0) = .Width * 0.05 - (30 * 3)
    .ColWidth(1) = .Width * 0.3 - (30 * 3)
    .ColWidth(2) = .Width * 0.65 - (40 * 3)
    .ColSel = 2
    .ColWidth(3) = 0
    .ColWidth(4) = 0
    .ColWidth(5) = 0
    .ColWidth(6) = 0
    .ColWidth(7) = 0
    .ColWidth(8) = 0
    .ColWidth(9) = 0
    .ColWidth(10) = 0
    .ColWidth(11) = 0
    .ColWidth(12) = 0
    .ColWidth(13) = 0
    .ColWidth(14) = 0
    .ColWidth(15) = 0
    .ColWidth(16) = 0
    .ColWidth(17) = 0
    .ColWidth(18) = 0
    .ColWidth(19) = 0
    .ColWidth(20) = 0
End With
End Sub

Private Sub assignValues(myRow As Integer)
With msfgResult
    txtPatientID.Text = .TextMatrix(myRow, 1)
    txtLastName.Text = .TextMatrix(myRow, 8)
    txtFirstName.Text = .TextMatrix(myRow, 9)
    txtMiddleName.Text = .TextMatrix(myRow, 10)
    txtBirthDate.Text = Format(.TextMatrix(myRow, 3), "mm/dd/yyyy")
    txtCivilStatus.Text = .TextMatrix(myRow, 5)
    txtAge.Text = .TextMatrix(myRow, 6)
    txtSex.Text = .TextMatrix(myRow, 7)
    txtSpecimen.Text = .TextMatrix(myRow, 4)
    txtPhysician.Text = .TextMatrix(myRow, 20)
    txtWard.Text = .TextMatrix(myRow, 12)
    txtRoom.Text = .TextMatrix(myRow, 13)
    If .TextMatrix(myRow, 18) = "<FNS>" Then
        txtRequestDate.Text = "<FNS>"
    Else
        txtRequestDate.Text = Format(.TextMatrix(myRow, 18), "mm/dd/yyyy") & " " & Format(.TextMatrix(myRow, 19), "HH:mm")
    End If
    If .TextMatrix(myRow, 15) = "<FNS>" Then
        txtResultdate.Text = "<FNS>"
    Else
        txtResultdate.Text = Format(.TextMatrix(myRow, 15), "mm/dd/yyyy") & " " & Format(.TextMatrix(myRow, 17), "HH:mm")
    End If
End With
End Sub

Private Sub ClearValues()
    txtPatientID.Text = ""
    txtLastName.Text = ""
    txtFirstName.Text = ""
    txtMiddleName.Text = ""
    txtBirthDate.Text = ""
    txtCivilStatus.Text = ""
    txtAge.Text = ""
    txtSex.Text = ""
    txtSpecimen.Text = ""
End Sub

Private Sub lblTitle_Change()
    Me.Caption = lblTitle.Caption
End Sub

Private Function GenerateTestPacket() As String
Dim tempStr As String
Dim mPort As String
Dim mAddress As String
Dim mPath As String
Dim tempPacket As String
Dim arrTemp() As String
Dim mData As String
Dim mServerName As String
Dim mFieldName As String
Dim mMethod As String
Dim mContentType As String
mServerName = GetSetting("LAB-I", "OBISettings", "ServerName", "")
mFieldName = GetSetting("LAB-I", "OBISettings", "RestApiTestField", "")
mMethod = GetSetting("LAB-I", "OBISettings", "RestApiMethod", "")
mContentType = GetSetting("LAB-I", "OBISettings", "RestApiContentType", "")
If Format(Left(Trim(mServerName), 7), ">") = "HTTP://" Then
    tempStr = Mid(Trim(mServerName), 8)
    mPath = Mid(tempStr, InStr(1, tempStr, "/"))
    tempStr = Mid(tempStr, 1, InStr(1, tempStr, "/") - 1)
Else
    tempStr = Mid(mServerName, 1, InStr(1, mServerName, "/") - 1)
    mPath = Mid(mServerName, InStr(1, mServerName, "/"))
End If
arrTemp = Split(tempStr, ":")
If InStr(1, tempStr, ":") > 0 Then
    mPort = arrTemp(1)
    mAddress = arrTemp(0)
Else
    mPort = "80"
    mAddress = arrTemp(0)
End If
tempPacket = ""
tempPacket = tempPacket & "Host: " & mAddress & vbCrLf
tempPacket = tempPacket & "Accept: */*" & vbCrLf
'tempPacket = tempPacket & "Expect: 100-Continue" & vbCrLf
tempPacket = tempPacket & "Content-Type: " & mContentType & vbCrLf

Select Case Format(mContentType, ">")
    Case "APPLICATION/JSON"
        mData = "{""" & mFieldName & """:""" & ySearchID & """}"
        If mMethod = "GET" Then
            If Len(Trim(mFieldName)) = 0 Then
                mData = ""
            Else
                mData = "/" & mFieldName & "/" & ySearchID
            End If
            tempPacket = mMethod & " " & mPath & mData & " HTTP/1.1" & vbCrLf & tempPacket & vbCrLf & vbCrLf
        Else
            tempPacket = mMethod & " " & mPath & " HTTP/1.1" & vbCrLf & tempPacket
            tempPacket = tempPacket & "Content-Length: " & Len(mData) & vbCrLf & vbCrLf & mData
        End If
    Case "APPLICATION/X-WWW-FORM-URLENCODED"
        mData = mFieldName & "=" & ySearchID
        tempPacket = mMethod & " " & mPath & "?" & mData & " HTTP/1.1" & vbCrLf & tempPacket
        If mMethod = "POST" Then
            tempPacket = tempPacket & "Content-Length: " & Len(mData) & vbCrLf & vbCrLf & mData
        End If
    Case Else
        tempPacket = mMethod & " " & mPath & " HTTP/1.1" & vbCrLf & tempPacket
        mData = ""
End Select
wsTest.Close
blnConnected = False
wsTest.Protocol = sckTCPProtocol
wsTest.RemoteHost = mAddress
wsTest.RemotePort = mPort
wsTest.Connect
TimeOutCtr = 0
tmrTimeOut.Enabled = True
While Not blnConnected And TimeOutCtr <= mTimeOut
    DoEvents
Wend
If blnConnected Then
    wsTest.SendData tempPacket
Else
    TimeOutCtr = 0
    tmrTimeOut.Enabled = False
    wsTest.Close
    blnConnected = False
    RequestDone = True
    DataReceived = "{""ERROR"":""Server cannot be found...""}"
    If Val(GetSetting("LAB-I", "OBISettings", "ShowErrorMessage", "0")) = 1 And GetError(DataReceived, 1) = "1" Then
        MsgBox GetError(DataReceived, 3), vbOKOnly + vbInformation, "PID Search"
    End If
End If
End Function

Private Sub tmrTimeOut_Timer()
TimeOutCtr = TimeOutCtr + 1
If TimeOutCtr > mTimeOut Then tmrTimeOut.Enabled = False
End Sub

Private Sub wsTest_Close()
    blnConnected = False
    wsTest.Close
    RequestDone = True
End Sub

Private Sub wsTest_Connect()
blnConnected = True
RequestDone = False
End Sub

Private Sub wsTest_DataArrival(ByVal bytesTotal As Long)
Dim strPageReturn As String
wsTest.GetData strPageReturn
DataReceived = ProcessReceived(strPageReturn)
RequestDone = True
End Sub

Private Sub wsTest_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    wsTest.Close
    RequestDone = True
End Sub

Private Function ProcessReceived(mess As String) As String
Dim mData As String
Dim p As Object
Dim sInputJson As String
Dim ctr As Integer
Dim myField As String
Dim myTable As String
myTable = GetSetting("LAB-I", "OBISettings", "RestApiTestTable", "")
If GetStatusCode(mess) = "200" Then
    If InStr(1, mess, vbCrLf & vbCrLf) > 0 Then
        mData = Mid(mess, InStr(1, mess, vbCrLf & vbCrLf) + 4)
    Else
        mData = mess
    End If
    
    mData = Replace(mData, """: [", """:[")
    If InStr(1, mData, """" & myTable & """:[", vbTextCompare) > 0 Then
        mData = Mid(mData, InStr(1, mData, """" & myTable & """:[", vbTextCompare) + Len("""" & myTable & """:["))
        mData = Mid(mData, 1, InStr(1, mData, "]"))
        mData = Mid(mData, 1, InStr(1, mData, "}"))
    End If
    mData = Replace(mData, """:NULL,", """:"""",", , , vbTextCompare)
    'Trim Carriage Return
    Do While Left(mData, 2) = vbCrLf
        mData = Mid(mData, 3)
    Loop
    Do While Right(mData, 2) = vbCrLf
        mData = Mid(mData, 1, Len(mData) - 2)
    Loop
    '--------------
    mData = Trim(mData)
    Set p = JSON.parse(mData)
       
    If Len(mData) > 0 Then
        With msfgResult
            .Clear
            .Cols = 21
            .rows = 1
            myField = GetSetting("LAB-I", "OBISettings", "FLD_1", "''")
            If myField <> "''" Then .TextMatrix(0, 1) = p.Item(myField) Else .TextMatrix(0, 1) = "<FNS>"
            myField = GetSetting("LAB-I", "OBISettings", "FLD_3", "''")
            If myField <> "''" Then .TextMatrix(0, 8) = p.Item(myField) Else .TextMatrix(0, 8) = "<FNS>"
            myField = GetSetting("LAB-I", "OBISettings", "FLD_4", "''")
            If myField <> "''" Then .TextMatrix(0, 9) = p.Item(myField) Else .TextMatrix(0, 9) = "<FNS>"
            myField = GetSetting("LAB-I", "OBISettings", "FLD_5", "''")
            If myField <> "''" Then .TextMatrix(0, 10) = p.Item(myField) Else .TextMatrix(0, 10) = "<FNS>"
            myField = GetSetting("LAB-I", "OBISettings", "FLD_17", "''")
            If myField <> "''" Then .TextMatrix(0, 3) = p.Item(myField) Else .TextMatrix(0, 3) = "<FNS>"
            myField = GetSetting("LAB-I", "OBISettings", "FLD_18", "''")
            If myField <> "''" Then .TextMatrix(0, 5) = p.Item(myField) Else .TextMatrix(0, 5) = "<FNS>"
            myField = GetSetting("LAB-I", "OBISettings", "FLD_19", "''")
            If myField <> "''" Then .TextMatrix(0, 4) = p.Item(myField) Else .TextMatrix(0, 4) = "<FNS>"
            myField = GetSetting("LAB-I", "OBISettings", "FLD_6", "''")
            If myField <> "''" Then .TextMatrix(0, 6) = p.Item(myField) Else .TextMatrix(0, 6) = "<FNS>"
            myField = GetSetting("LAB-I", "OBISettings", "FLD_7", "''")
            If myField <> "''" Then .TextMatrix(0, 7) = p.Item(myField) Else .TextMatrix(0, 7) = "<FNS>"
            myField = GetSetting("LAB-I", "OBISettings", "FLD_3", "''")
            If myField <> "''" Then .TextMatrix(0, 8) = p.Item(myField) Else .TextMatrix(0, 8) = "<FNS>"
            myField = GetSetting("LAB-I", "OBISettings", "FLD_4", "''")
            If myField <> "''" Then .TextMatrix(0, 9) = p.Item(myField) Else .TextMatrix(0, 9) = "<FNS>"
            myField = GetSetting("LAB-I", "OBISettings", "FLD_5", "''")
            If myField <> "''" Then .TextMatrix(0, 10) = p.Item(myField) Else .TextMatrix(0, 10) = "<FNS>"
            myField = GetSetting("LAB-I", "OBISettings", "FLD_8", "''")
            If myField <> "''" Then .TextMatrix(0, 12) = p.Item(myField) Else .TextMatrix(0, 12) = "<FNS>"
            myField = GetSetting("LAB-I", "OBISettings", "FLD_9", "''")
            If myField <> "''" Then .TextMatrix(0, 13) = p.Item(myField) Else .TextMatrix(0, 13) = "<FNS>"
            
        
            myField = GetSetting("LAB-I", "OBISettings", "FLD_13", "''")
            If myField <> "''" Then .TextMatrix(0, 18) = p.Item(myField) Else .TextMatrix(0, 18) = "<FNS>"
            myField = GetSetting("LAB-I", "OBISettings", "FLD_14", "''")
            If myField <> "''" Then .TextMatrix(0, 19) = p.Item(myField) Else .TextMatrix(0, 19) = "<FNS>"
            myField = GetSetting("LAB-I", "OBISettings", "FLD_12", "''")
            If myField <> "''" Then .TextMatrix(0, 15) = p.Item(myField) Else .TextMatrix(0, 15) = "<FNS>"
            
    '        myField = GetSetting("LAB-I", "OBISettings", "FLD_12", "''")
    '        If Len(Trim(myField)) > 0 Then .TextMatrix(0, 18) = p.Item(myField)
    '        myField = GetSetting("LAB-I", "OBISettings", "FLD_13", "''")
    '        If Len(Trim(myField)) > 0 Then .TextMatrix(0, 19) = p.Item(myField)
    '        myField = GetSetting("LAB-I", "OBISettings", "FLD_14", "''")
    '        If Len(Trim(myField)) > 0 Then .TextMatrix(0, 20) = p.Item(myField)
        End With
    End If
    assignValues 0
Else
    mData = "{""ERROR"":""" & GetStatusCode(mess) & ": " & GetHTTPStatusCodeDesc(Val(GetStatusCode(mess)), False) & """}"
End If
If Val(GetSetting("LAB-I", "OBISettings", "ShowErrorMessage", "0")) = 1 And GetError(mData, 1) = "1" Then
    MsgBox GetError(mData, 3), vbOKOnly + vbInformation, "PID Search"
End If
ProcessReceived = mData
End Function

Private Function GetStatusCode(mess As String) As String
Dim arrTemp() As String
Dim tempStr As String
tempStr = ""
If Len(mess) > 0 And InStr(1, mess, " ") Then
    tempStr = Mid(mess, 1, InStr(1, mess, vbCrLf))
    tempStr = Replace(tempStr, "   ", " ")
    tempStr = Replace(tempStr, "  ", " ")
    arrTemp = Split(tempStr, " ")
    If UBound(arrTemp) >= 1 Then
        tempStr = arrTemp(1)
    End If
End If
GetStatusCode = tempStr
End Function

Private Function GetHTTPStatusCodeDesc(ByVal pintHTTPStatusCode As Integer, pbolVerbose As Boolean) As String
    '*******************************************************************************
    '
    ' Function       : GetHTTPStatusCodeDesc
    '
    ' Return Type    : String
    '
    ' Copyright      : 2004 - Salty Brine Software (Art Araya)
    '
    ' Description    : Provides either a brief or verbose description for an
    '                  HTTP Status Code.
    '
    ' Parameters     : 1) pintHTTPStatusCode (Integer) - the HTTP Status Code
    '                     that a desciption Is needed For
    '                : 2) pbolVerbose (Boolean) - Used to specify whether a brief or
    '                     verbose description of the status code Is desired.
    '
    ' Range Description
    ' 100-199 Informational
    ' 200-299 Success: indicates that a client's request arrived and was accepted.
    ' 300-399 Client request redirection, further action required.
    ' 400-499 Client error
    ' 500-599 Server error
    '
    '*******************************************************************************

    Dim strHTTPStatusCodeDesc   As String

    If pbolVerbose Then

        Select Case pintHTTPStatusCode
            Case 100
                strHTTPStatusCodeDesc = "Continue - The client can continue its request."
            Case 101
                strHTTPStatusCodeDesc = "Switching Protocols - The server has complied with the client's request to switch protocols."
            Case 200
                strHTTPStatusCodeDesc = "OK - The request was successful."
            Case 201
                strHTTPStatusCodeDesc = "Created - The request was successful and a new resource was created as a result."
            Case 202
                strHTTPStatusCodeDesc = "Accepted - The request was accepted for processing."
            Case 203
                strHTTPStatusCodeDesc = "Non-Authoritative Information - The meta-information in the entity-header is from a local or third-party copy."
            Case 204
                strHTTPStatusCodeDesc = "No Content - The server serviced the request but there is no information to return."
            Case 205
                strHTTPStatusCodeDesc = "Reset Content - The request was successful and the user agent should clear the input form for further input."
            Case 206
                strHTTPStatusCodeDesc = "Partial Content - The server serviced a (byte) range request for the resource."
            Case 300
                strHTTPStatusCodeDesc = "Multiple Choices - The requested resource could be one of multiple resources."
            Case 301
                strHTTPStatusCodeDesc = "Moved Permanently - The requested resource has permanently moved to a new location."
            Case 302
                strHTTPStatusCodeDesc = "Moved Temporarily - The requested resource has temporarily moved to a new location."
            Case 303
                strHTTPStatusCodeDesc = "See Other - The response to the request is under a different URI and can be retrieved with a GET request."
            Case 304
                strHTTPStatusCodeDesc = "Not Modified - The requested resource has not been modified since it was last requested."
            Case 305
                strHTTPStatusCodeDesc = "Use Proxy - The requested resource must be accessed through a proxy server."
            Case 400
                strHTTPStatusCodeDesc = "Bad Request - The server cannot read the request because its syntax is incorrect."
            Case 401
                strHTTPStatusCodeDesc = "Unauthorized - The server must authenticate the user before servicing the request."
            Case 402
                strHTTPStatusCodeDesc = "Payment Required"
            Case 403
                strHTTPStatusCodeDesc = "Forbidden - The server refused to service the request."
            Case 404
                strHTTPStatusCodeDesc = "Not Found - The server cannot find the requested resource."
            Case 405
                strHTTPStatusCodeDesc = "Method Not Allowed - The method specified in the Request-Line is not permitted for the requested resource."
            Case 406
                strHTTPStatusCodeDesc = "Not Acceptable - The requested resource can only generate response entities that have unacceptable content characteristics according to the accept headers sent in the request."
            Case 407
                strHTTPStatusCodeDesc = "Proxy Authentication Required - The proxy server must authenticate the user before servicing the request."
            Case 408
                strHTTPStatusCodeDesc = "Request Timeout - The client did not make its request within the amount of time the server will wait for requests."
            Case 409
                strHTTPStatusCodeDesc = "Conflict - The server could not service the request due to a current conflict with the requested resource."
            Case 410
                strHTTPStatusCodeDesc = "Gone - The requested resource is no longer available on the server."
            Case 411
                strHTTPStatusCodeDesc = "Length Required - The server will not service the request without a Content-Length specified in the request."
            Case 412
                strHTTPStatusCodeDesc = "Precondition Failed - A precondition specified in one or more of the request-header fields failed."
            Case 413
                strHTTPStatusCodeDesc = "Request Entity Too Large - The server will not service the request because the requested resource is too large."
            Case 414
                strHTTPStatusCodeDesc = "Request-URI Too Long - The server will not service the request because the requested URL is too long."
            Case 415
                strHTTPStatusCodeDesc = "Unsupported Media Type - The server will not service the request because the format of the request is not supported by the requested resource for the requested method."
            Case 500
                strHTTPStatusCodeDesc = "Internal Server Error - The server could not service the request because of an unexpected internal error."
            Case 501
                strHTTPStatusCodeDesc = "Not Implemented - The sever cannot service the request because it does not support the request method."
            Case 502
                strHTTPStatusCodeDesc = "Bad Gateway - The proxy server received an invalid response from the content server or another proxy server in a proxy chain."
            Case 503
                strHTTPStatusCodeDesc = "Service Unavailable - The server could not service the request because it was temporarily overloaded or undergoing maintenance."
            Case 504
                strHTTPStatusCodeDesc = "Gateway Timeout - The proxy server did not receive a response from a chained proxy server or the origin content server within an acceptable amount of time."
            Case 505
                strHTTPStatusCodeDesc = "HTTP Version Not Supported - The server does not support the HTTP version specified in the request."
            Case Else
                strHTTPStatusCodeDesc = "Unrecognized HTTP Status Code"
        End Select

    Else

        Select Case pintHTTPStatusCode
            Case 100
                strHTTPStatusCodeDesc = "Continue"
            Case 101
                strHTTPStatusCodeDesc = "Switching Protocols"
            Case 200
                strHTTPStatusCodeDesc = "OK"
            Case 201
                strHTTPStatusCodeDesc = "Created"
            Case 202
                strHTTPStatusCodeDesc = "Accepted"
            Case 203
                strHTTPStatusCodeDesc = "Non-Authoritative Information"
            Case 204
                strHTTPStatusCodeDesc = "No Content"
            Case 205
                strHTTPStatusCodeDesc = "Reset Content"
            Case 206
                strHTTPStatusCodeDesc = "Partial Content"
            Case 300
                strHTTPStatusCodeDesc = "Multiple Choices"
            Case 301
                strHTTPStatusCodeDesc = "Moved Permanently"
            Case 302
                strHTTPStatusCodeDesc = "Moved Temporarily"
            Case 303
                strHTTPStatusCodeDesc = "See Other"
            Case 304
                strHTTPStatusCodeDesc = "Not Modified"
            Case 305
                strHTTPStatusCodeDesc = "Use Proxy"
            Case 400
                strHTTPStatusCodeDesc = "Bad Request"
            Case 401
                strHTTPStatusCodeDesc = "Unauthorized"
            Case 402
                strHTTPStatusCodeDesc = "Payment Required"
            Case 403
                strHTTPStatusCodeDesc = "Forbidden"
            Case 404
                strHTTPStatusCodeDesc = "Not Found"
            Case 405
                strHTTPStatusCodeDesc = "Method Not Allowed"
            Case 406
                strHTTPStatusCodeDesc = "Not Acceptable"
            Case 407
                strHTTPStatusCodeDesc = "Proxy Authentication Required"
            Case 408
                strHTTPStatusCodeDesc = "Request Timeout"
            Case 409
                strHTTPStatusCodeDesc = "Conflict"
            Case 410
                strHTTPStatusCodeDesc = "Gone"
            Case 411
                strHTTPStatusCodeDesc = "Length Required"
            Case 412
                strHTTPStatusCodeDesc = "Precondition Failed"
            Case 413
                strHTTPStatusCodeDesc = "Request Entity Too Large"
            Case 414
                strHTTPStatusCodeDesc = "Request-URI Too Long"
            Case 415
                strHTTPStatusCodeDesc = "Unsupported Media Type"
            Case 500
                strHTTPStatusCodeDesc = "Internal Server Error"
            Case 501
                strHTTPStatusCodeDesc = "Not Implemented"
            Case 502
                strHTTPStatusCodeDesc = "Bad Gateway"
            Case 503
                strHTTPStatusCodeDesc = "Service Unavailable"
            Case 504
                strHTTPStatusCodeDesc = "Gateway Timeout"
            Case 505
                strHTTPStatusCodeDesc = "HTTP Version Not Supported"
            Case Else
                strHTTPStatusCodeDesc = "Unrecognized HTTP Status Code"
        End Select
    End If  'If pbolVerbose

    GetHTTPStatusCodeDesc = strHTTPStatusCodeDesc

End Function


Private Function GetError(mess As String, mTrans As Integer) As String
Dim arrTemp() As String
Dim tempStr As String
Dim hasError As String
tempStr = Replace(mess, "{""", "")
tempStr = Replace(tempStr, """}", "")
tempStr = Replace(tempStr, """:""", "|")
arrTemp = Split(tempStr, "|")
hasError = "0"
If UBound(arrTemp) >= 1 Then
    If arrTemp(0) = "ERROR" Then
        hasError = "1"
    End If
End If
If hasError = "1" Then
    Select Case mTrans
        Case 1: GetError = hasError
        Case 2: GetError = arrTemp(0)
        Case 3: GetError = arrTemp(1)
    End Select
Else
    Select Case mTrans
        Case 1: GetError = hasError
        Case 2: GetError = ""
        Case 3: GetError = ""
    End Select
End If
End Function

