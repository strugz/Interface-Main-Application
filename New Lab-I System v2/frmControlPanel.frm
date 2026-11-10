VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmControlPanel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Program Setup"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   1380
      Width           =   6135
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
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
         Left            =   5850
         MouseIcon       =   "frmControlPanel.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   210
         Width           =   135
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Program Setup"
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
         Width           =   1455
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   5760
         Top             =   180
         Width           =   285
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   4035
      Index           =   0
      Left            =   7110
      TabIndex        =   2
      Top             =   2100
      Width           =   1875
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlPanel.frx":030A
            Key             =   "formsettings"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlPanel.frx":11E4
            Key             =   "medtech"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlPanel.frx":20BE
            Key             =   "nvclass"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlPanel.frx":2758
            Key             =   "nvitems"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlPanel.frx":2DF2
            Key             =   "database"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlPanel.frx":3CCC
            Key             =   "pathologist"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlPanel.frx":4BA6
            Key             =   "physician"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlPanel.frx":5A80
            Key             =   "profiles"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlPanel.frx":611A
            Key             =   "profileexams"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlPanel.frx":67B4
            Key             =   "programsettings"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlPanel.frx":6E4E
            Key             =   "reportformat"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlPanel.frx":7D28
            Key             =   "selectionlists"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlPanel.frx":8C02
            Key             =   "labexams"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlPanel.frx":9ADC
            Key             =   "flaglists"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlPanel.frx":A9B6
            Key             =   "dbtools"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlPanel.frx":B890
            Key             =   "comset"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlPanel.frx":C76A
            Key             =   "hospital"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   5835
      Left            =   60
      TabIndex        =   4
      Top             =   30
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   10292
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
setFormColors Me
With lvwItems
    .ListItems.Add 1, "programsettings", "01 - Program Settings", "programsettings"
    .ListItems.Add 1, "database", "04 - Database Manager", "dbtools"
    .ListItems.Add 1, "nvclass", "15 - Normal Values Class", "nvclass"
    .ListItems.Add 1, "nvitems", "16 - Normal Values Items", "nvitems"
    .ListItems.Add 1, "profiles", "13 - Profiles", "profiles"
    .ListItems.Add 1, "profileexams", "14 - Profile Exams", "profileexams"
    .ListItems.Add 1, "formsettings", "02 - Result Form Settings", "formsettings"
    If SectionCode = "PULMO" Then
        .ListItems.Add 1, "medtech", "09 - Respiratory Therapist", "medtech"
        .ListItems.Add 1, "pathologist", "11 - Pulmonologist", "pathologist"
    Else
        .ListItems.Add 1, "medtech", "09 - Medical Technologists", "medtech"
        .ListItems.Add 1, "pathologist", "11 - Pathologists", "pathologist"
    End If
    .ListItems.Add 1, "physician", "10 - Physicians", "physician"
    .ListItems.Add 1, "flaglists", "07 - Flag Lists", "flaglists"
    .ListItems.Add 1, "selectionlists", "06 - Selection Lists Update", "selectionlists"
    .ListItems.Add 1, "comset", "03 - Communication Settings", "comset"
    .ListItems.Add 1, "labexam", "12 - Laboratory Exams", "labexams"
    .ListItems.Add 1, "reportformats", "05 - Report Formats", "reportformat"
    .ListItems.Add 1, "hospitalprofile", "08 - Hospital Profile", "hospital"
    .ListItems.Add 1, "middleware", "17 - Middleware (LIS/HIS)", "comset"
    If isConnected Then
        .ListItems("comset").Ghosted = True
        .ListItems("labexam").Ghosted = True
    Else
        .ListItems("comset").Ghosted = False
        .ListItems("labexam").Ghosted = False
    End If
End With
End Sub

Private Sub Label2_Click()
    Unload Me
End Sub

Private Sub lvwItems_DblClick()
Dim EnablePlugin As Integer
Dim PluginName As String
If Not lvwItems.SelectedItem.Ghosted Then
    Select Case lvwItems.SelectedItem.Key
        Case "programsettings"
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
        Case "formsettings"
            HostSA.showformsettings
        Case "hospitalprofile"
            HostSA.showHospitalInfo
        Case "labexam"
            HostSA.showLabExams
        Case "labprofile"
            HostSA.showLabProfiles
        Case "medtech"
            HostSA.showmedTech
        Case "nvitems"
            HostSA.showNormalValues
        Case "pathologist"
            HostSA.showPathologist
        Case "nvclass"
            HostSA.showClassifications
        Case "profiles"
            HostSA.showLabProfiles
        Case "profileexams"
            HostSA.showProfileExams
        Case "physician"
            HostSA.showphysician
        Case "flaglists"
            HostSA.showRemarksTemplate
        Case "labexam"
            HostSA.showLabExams
        Case "database"
            If isResultVisible Then frmSearchResult.Visible = False
            frmDbMan.Show 1
        Case "comset"
            mComSet.showsettings
        Case "selectionlists"
            HostSA.showSelectionLists
        Case "middleware"
            Unload Me                       'a mode switch restarts Lab-I, so close this panel first
            mdiMain.ShowMiddlewareSettings
        Case "reportformats"
            HostSA.showReportFormats
            setResultReport
            HostSA.wRPTname = RPTName
    End Select
End If
End Sub
    
Private Sub lvwItems_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
ElseIf KeyCode = vbKeyReturn Then
    lvwItems_DblClick
End If
End Sub
