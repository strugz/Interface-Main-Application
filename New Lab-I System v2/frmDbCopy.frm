VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDbCopy 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2370
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   8475
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database Copy"
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
         Width           =   1470
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   2055
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   300
      Width           =   8475
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   165
         Left            =   180
         TabIndex        =   9
         Top             =   1260
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   0
         Min             =   1
         Max             =   36
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3810
         TabIndex        =   8
         Top             =   570
         Width           =   4575
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   540
         Width           =   2085
      End
      Begin VB.TextBox txtDatabase 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   6705
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
         Height          =   345
         Left            =   7050
         TabIndex        =   3
         Top             =   1590
         Width           =   1305
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   510
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   8235
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database Password:"
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
         Left            =   120
         TabIndex        =   6
         Top             =   630
         Width           =   1530
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database Location:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   210
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmDBcopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdStart_Click()
    StartCopy
End Sub

Private Sub Form_Load()
    setFormColors Me
    txtDatabase.Text = App.Path & "\Previous\Data\"
End Sub

Private Sub StartCopy()
Dim TempLocation As String
If MsgBox("Are you sure you want to copy old data?", vbYesNo + vbQuestion + vbDefaultButton2, "Data Copy") = vbYes Then
    DoEvents
    lblStatus.Caption = "Setting up location..."
    pBar.Value = 1
    lblStatus.Refresh
    TempLocation = Replace(Format(DBDatabase, ">"), "LABORATORY.MDB", "")
    lblStatus.Caption = "Loading Patient Data..."
    pBar.Value = 2
    lblStatus.Refresh
    LinkTable "tbPatientMaster", txtDatabase.Text & "LabPatientData.mdb", DBDatabase, "tempMaster", txtPassword.Text
    lblStatus.Caption = "Loading Patient Results..."
    pBar.Value = 3
    lblStatus.Refresh
    LinkTable "tbPatientValues", txtDatabase.Text & "LabPatientValues.mdb", DBDatabase, "tempValues", txtPassword.Text
    lblStatus.Caption = "Loading Instrument Results..."
    pBar.Value = 4
    lblStatus.Refresh
    LinkTable "tbLabInstrumentResults", txtDatabase.Text & "InstrumentResults.mdb", DBDatabase, "tempResults", txtPassword.Text
    lblStatus.Caption = "Loading Exams..."
    pBar.Value = 5
    lblStatus.Refresh
    LinkTable "tbCoLabExam", txtDatabase.Text & "LABORATORY.MDB", DBDatabase, "tempExam", txtPassword.Text
    lblStatus.Caption = "Loading Class..."
    pBar.Value = 6
    lblStatus.Refresh
    LinkTable "tbLabClassifications", txtDatabase.Text & "LABORATORY.MDB", DBDatabase, "tempClass", txtPassword.Text
    lblStatus.Caption = "Loading Class Values..."
    pBar.Value = 7
    lblStatus.Refresh
    LinkTable "tbLabClassExams", txtDatabase.Text & "LABORATORY.MDB", DBDatabase, "tempClassVal", txtPassword.Text
    lblStatus.Caption = "Loading Profiles..."
    pBar.Value = 8
    lblStatus.Refresh
    LinkTable "tbLabProfiles", txtDatabase.Text & "LABORATORY.MDB", DBDatabase, "tempProfile", txtPassword.Text
    lblStatus.Caption = "Loading Profile Exams..."
    pBar.Value = 9
    lblStatus.Refresh
    LinkTable "tbLabProfileExams", txtDatabase.Text & "LABORATORY.MDB", DBDatabase, "tempProfileVal", txtPassword.Text
    lblStatus.Caption = "Loading Complete..."
    pBar.Value = 10
    lblStatus.Refresh
    lblStatus.Caption = "Clean-up Patient Data"
    lblStatus.Refresh
    myConnection.BeginTrans
    pBar.Value = 11
    lblStatus.Refresh
    myConnection.Execute "Delete from tbpatientmaster"
    lblStatus.Caption = "Clean-up Patient Results..."
    pBar.Value = 12
    lblStatus.Refresh
    myConnection.Execute "delete from tbPatientValues"
    lblStatus.Caption = "Clean-up Instrument Results..."
    pBar.Value = 13
    lblStatus.Refresh
    myConnection.Execute "delete from tbLabInstrumentResults"
    lblStatus.Caption = "Clean-up Exams..."
    pBar.Value = 14
    lblStatus.Refresh
    myConnection.Execute "delete from tbCoLabExam"
    lblStatus.Caption = "Clean-up Class..."
    pBar.Value = 15
    lblStatus.Refresh
    myConnection.Execute "delete from tbLabClassifications"
    lblStatus.Caption = "Clean-up Class Values..."
    pBar.Value = 16
    lblStatus.Refresh
    myConnection.Execute "delete from tbLabClassExams"
    lblStatus.Caption = "Clean-up Profiles..."
    pBar.Value = 17
    lblStatus.Refresh
    myConnection.Execute "delete from tbLabProfiles"
    lblStatus.Caption = "Clean-up Profile Exams..."
    pBar.Value = 18
    lblStatus.Refresh
    myConnection.Execute "delete from tbLabProfileExams"
    lblStatus.Caption = "Clean-up Complete..."
    pBar.Value = 19
    lblStatus.Refresh
    
    pBar.Value = 20
    lblStatus.Refresh
    myConnection.Execute "insert into tbpatientmaster select * from tempMaster"
    lblStatus.Caption = "Uploading Patient Results..."
    pBar.Value = 21
    lblStatus.Refresh
    myConnection.Execute "insert into tbPatientValues select * from tempValues"
    lblStatus.Caption = "Uploading Instrument Results..."
    pBar.Value = 22
    lblStatus.Refresh
    myConnection.Execute "insert into tbLabInstrumentResults select * from tempResults"
    lblStatus.Caption = "Uploading Exams..."
    pBar.Value = 23
    lblStatus.Refresh
    myConnection.Execute "insert into tbCoLabExam select * from tempExam"
    lblStatus.Caption = "Uploading Class..."
    pBar.Value = 24
    lblStatus.Refresh
    myConnection.Execute "insert into tbLabClassifications select * from tempClass"
    lblStatus.Caption = "Uploading Class Values..."
    pBar.Value = 25
    lblStatus.Refresh
    myConnection.Execute "insert into tbLabClassExams select * from tempClassVal"
    lblStatus.Caption = "Uploading Profiles..."
    pBar.Value = 26
    lblStatus.Refresh
    myConnection.Execute "insert into tbLabProfiles select * from tempProfile"
    lblStatus.Caption = "Uploading Profile Exams..."
    pBar.Value = 27
    lblStatus.Refresh
    myConnection.Execute "insert into tbLabProfileExams select * from tempProfileVal"
    lblStatus.Caption = "Uploading Complete..."
    pBar.Value = 28
    lblStatus.Refresh

    myConnection.Execute "drop table tempMaster"
    lblStatus.Caption = "Removing Temporary Tables"
    pBar.Value = 29
    lblStatus.Refresh
    myConnection.Execute "drop table tempValues"
    lblStatus.Caption = "Removing Temporary Tables"
    pBar.Value = 30
    lblStatus.Refresh
    myConnection.Execute "drop table tempResults"
    lblStatus.Caption = "Removing Temporary Tables"
    pBar.Value = 31
    lblStatus.Refresh
    myConnection.Execute "drop table tempExam"
    lblStatus.Caption = "Removing Temporary Tables"
    pBar.Value = 32
    lblStatus.Refresh
    myConnection.Execute "drop table tempClass"
    lblStatus.Caption = "Removing Temporary Tables"
    pBar.Value = 33
    lblStatus.Refresh
    myConnection.Execute "drop table tempClassVal"
    lblStatus.Caption = "Removing Temporary Tables"
    pBar.Value = 34
    lblStatus.Refresh
    myConnection.Execute "drop table tempProfile"
    lblStatus.Caption = "Removing Temporary Tables"
    pBar.Value = 35
    lblStatus.Refresh
    myConnection.Execute "drop table tempProfileVal"
    lblStatus.Caption = "Done"
    pBar.Value = 36
    lblStatus.Refresh
    myConnection.CommitTrans
End If
End Sub
