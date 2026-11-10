VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataBaseCopy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Copy"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8610
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8610
   StartUpPosition =   1  'CenterOwner
   Begin VB.DriveListBox Drive1 
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
      Height          =   330
      Left            =   1590
      TabIndex        =   12
      Top             =   2580
      Width           =   6735
   End
   Begin VB.DirListBox Dir1 
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
      Height          =   2010
      Left            =   1590
      TabIndex        =   11
      Top             =   540
      Width           =   6735
   End
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   9210
      TabIndex        =   0
      Top             =   2430
      Width           =   1725
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
      Height          =   6105
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   8475
      Begin VB.TextBox txtRecords 
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
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   4320
         Width           =   1995
      End
      Begin VB.TextBox txtInstitution 
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
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   3420
         Width           =   7365
      End
      Begin VB.TextBox txtMinor 
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
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   3870
         Width           =   585
      End
      Begin VB.TextBox txtMajor 
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
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   3870
         Width           =   555
      End
      Begin VB.CommandButton cmdCheck 
         Caption         =   "Verify Database"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1500
         TabIndex        =   17
         Top             =   2910
         Width           =   6735
      End
      Begin VB.TextBox txtHeaderCode 
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
         Left            =   3780
         MaxLength       =   1
         TabIndex        =   16
         Top             =   3870
         Width           =   705
      End
      Begin MSComctlLib.ProgressBar pbResultID 
         Height          =   105
         Left            =   150
         TabIndex        =   13
         Top             =   5190
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   185
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   165
         Left            =   150
         TabIndex        =   9
         Top             =   5010
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   0
         Min             =   1
         Max             =   52
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5430
         TabIndex        =   8
         Top             =   5460
         Width           =   1425
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
         Left            =   6180
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   3840
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
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   6735
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
         Height          =   435
         Left            =   6930
         TabIndex        =   3
         Top             =   5460
         Width           =   1305
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "** Use Version 2.1.74 if you want to copy version lower than 2.1.55"
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
         Left            =   120
         TabIndex        =   26
         Top             =   5610
         Width           =   4920
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Records:"
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
         Left            =   150
         TabIndex        =   24
         Top             =   4380
         Width           =   660
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Institution:"
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
         Left            =   150
         TabIndex        =   22
         Top             =   3480
         Width           =   720
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rev:"
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
         Left            =   1410
         TabIndex        =   21
         Top             =   3930
         Width           =   330
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version:"
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
         Left            =   150
         TabIndex        =   18
         Top             =   3930
         Width           =   615
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Header Code:"
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
         Left            =   2700
         TabIndex        =   15
         Top             =   3930
         Width           =   990
      End
      Begin VB.Label lblText 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Copying..."
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
         Left            =   3390
         TabIndex        =   14
         Top             =   4770
         Visible         =   0   'False
         Width           =   4785
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
         Height          =   600
         Left            =   60
         TabIndex        =   10
         Top             =   4770
         Width           =   8145
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
         Left            =   4620
         TabIndex        =   6
         Top             =   3930
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
Attribute VB_Name = "frmDataBaseCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheck_Click()
Dim conTemp As New ADODB.Connection
Dim rstTemp As ADODB.Recordset
On Error GoTo errFound
    conTemp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtDatabase & "Laboratory.mdb;User Id=;Password=;Jet OLEDB:Database Password=" & IIf(Len(Trim(txtPassword.Text)) > 0, txtPassword, DBPassword) & ";"
    conTemp.ConnectionTimeout = 30
    conTemp.CursorLocation = adUseClient
    conTemp.Open
    Set rstTemp = conTemp.Execute("Select * from tbversion")
    txtMajor.Text = rstTemp.Fields("Major").Value
    txtMinor.Text = rstTemp.Fields("Minor").Value
    Set rstTemp = conTemp.Execute("Select * from tbHospitalInfo")
    txtInstitution.Text = rstTemp.Fields("HospitalName").Value & ""
    rstTemp.Close
    conTemp.Close
    conTemp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtDatabase & "InstrumentResults.mdb;User Id=;Password=;Jet OLEDB:Database Password=" & IIf(Len(Trim(txtPassword.Text)) > 0, txtPassword, DBPassword) & ";"
    conTemp.ConnectionTimeout = 30
    conTemp.CursorLocation = adUseClient
    conTemp.Open
'    Set rstTemp = conTemp.Execute("Select distinct min(sortnum) as HeaderCode from tbLabInstrumentResults where instrumentid ='" & InstrumentID & "'")
'    txtHeaderCode.Text = rstTemp.Fields("HeaderCode").Value
    Set rstTemp = conTemp.Execute("Select * from tbLabInstrumentResults where instr(1,refnum,'N') > 0 or Sortnum = " & txtHeaderCode.Text & "")
    txtRecords.Text = rstTemp.RecordCount
errFound:
    If Err.Number <> 0 Then
        Select Case Err.Number
            Case -2147467259
                MsgBox "Unable to connect to server" & vbCrLf & vbCrLf & "Database not found...", vbExclamation + vbOKOnly, App.ProductName
            Case -2147217865
                MsgBox "Old Database detected." & vbCrLf & vbCrLf & "Please Use Previous Version...", vbExclamation + vbOKOnly, App.ProductName
            Case Else
                MsgBox Err.Description & vbCrLf & "Please Contact your Systems Administrator.", vbExclamation + vbOKOnly, App.ProductName
        End Select
    Else
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdStart_Click()
    If Len(Trim(txtDatabase.Text)) > 0 Then
        StartCopy
    Else
        MsgBox "Invalid Directory", vbInformation + vbOKOnly
    End If
End Sub

Private Sub Dir1_Change()
    If Right(Dir1.Path, 1) = "\" Then
        txtDatabase.Text = Dir1.Path
    Else
        txtDatabase.Text = Dir1.Path & "\"
    End If
    txtInstitution.Text = ""
    txtMajor.Text = ""
    txtMinor.Text = ""
    'txtHeaderCode.Text = "0"
    txtRecords = ""
    cmdStart.Enabled = False
End Sub

Private Sub Dir1_Click()
    Dir1_Change
End Sub

Private Sub Drive1_Change()
On Error GoTo errFound
Dir1.Path = Drive1.Drive
errFound:
    If Err.Number > 0 Then
        MsgBox Err.Description
        Dir1.Visible = False
        txtDatabase.Text = ""
    Else
        Dir1.Visible = True
    End If
End Sub

Private Sub Form_Load()
    setFormColors Me
    txtDatabase.Text = App.Path & "\Previous\Data\"
    Dir1.Path = App.Path
    txtHeaderCode.Text = "0"
End Sub

Private Sub StartCopy()
Dim TempLocation As String
If Len(Trim(Dir(txtDatabase.Text & "LABORATORY.MDB"))) = 0 Then
    MsgBox "Database not found.  Please select another location.", vbInformation + vbOKOnly
    Exit Sub
Else
    Me.MousePointer = 11
    If MsgBox("Are you sure you want to copy old data?", vbYesNo + vbQuestion + vbDefaultButton2, "Data Copy") = vbYes Then
        DoEvents
        lblStatus.Caption = "Setting up location..."
        pBar.Value = 1
        lblStatus.Refresh
        TempLocation = Replace(Format(DBDatabase, ">"), "LABORATORY.MDB", "")
        lblStatus.Caption = "Connecting Patient Data..."
        pBar.Value = 2
        lblStatus.Refresh
        LinkTable "tbPatientMaster", txtDatabase.Text & "LabPatientData.mdb", DBDatabase, "tempMaster", txtPassword.Text
        lblStatus.Caption = "Connecting Patient Results..."
        pBar.Value = 3
        lblStatus.Refresh
        LinkTable "tbPatientValues", txtDatabase.Text & "LabPatientValues.mdb", DBDatabase, "tempValues", txtPassword.Text
        lblStatus.Caption = "Connecting Instrument Results..."
        pBar.Value = 4
        lblStatus.Refresh
        LinkTable "tbLabInstrumentResults", txtDatabase.Text & "InstrumentResults.mdb", DBDatabase, "tempResults", txtPassword.Text
        lblStatus.Caption = "Connecting Exams..."
        pBar.Value = 5
        lblStatus.Refresh
        LinkTable "tbCoLabExam", txtDatabase.Text & "LABORATORY.MDB", DBDatabase, "tempExam", txtPassword.Text
        lblStatus.Caption = "Connecting Class..."
        pBar.Value = 6
        lblStatus.Refresh
        LinkTable "tbLabClassifications", txtDatabase.Text & "LABORATORY.MDB", DBDatabase, "tempClass", txtPassword.Text
        lblStatus.Caption = "Connecting Class Values..."
        pBar.Value = 7
        lblStatus.Refresh
        LinkTable "tbLabClassExams", txtDatabase.Text & "LABORATORY.MDB", DBDatabase, "tempClassVal", txtPassword.Text
        lblStatus.Caption = "Connecting Profiles..."
        pBar.Value = 8
        lblStatus.Refresh
        LinkTable "tbLabProfiles", txtDatabase.Text & "LABORATORY.MDB", DBDatabase, "tempProfile", txtPassword.Text
        lblStatus.Caption = "Connecting Profile Exams..."
        pBar.Value = 9
        lblStatus.Refresh
        LinkTable "tbLabProfileExams", txtDatabase.Text & "LABORATORY.MDB", DBDatabase, "tempProfileVal", txtPassword.Text
        lblStatus.Caption = "Connecting Hospital Information..."
        pBar.Value = 10
        lblStatus.Refresh
        LinkTable "tbHospitalInfo", txtDatabase.Text & "LABORATORY.MDB", DBDatabase, "tempHospital", txtPassword.Text
        lblStatus.Caption = "Connecting Pathologist..."
        pBar.Value = 11
        lblStatus.Refresh
        LinkTable "tbLabPathologists", txtDatabase.Text & "LABORATORY.MDB", DBDatabase, "tempPathologists", txtPassword.Text
        lblStatus.Caption = "Connecting Physician..."
        pBar.Value = 12
        lblStatus.Refresh
        LinkTable "tbPhysician", txtDatabase.Text & "LABORATORY.MDB", DBDatabase, "tempPhysician", txtPassword.Text
        lblStatus.Caption = "Connecting Med.Tech..."
        pBar.Value = 13
        lblStatus.Refresh
        LinkTable "tbMedtech", txtDatabase.Text & "LABORATORY.MDB", DBDatabase, "tempMedTech", txtPassword.Text
        lblStatus.Caption = "Connecting Complete..."
        pBar.Value = 14
        lblStatus.Refresh
        

        lblStatus.Caption = "Clean-up Patient Data"
        lblStatus.Refresh
        myConnection.BeginTrans
        pBar.Value = 15
        lblStatus.Refresh
        myConnection.Execute "Delete from tbpatientmaster"
        lblStatus.Caption = "Clean-up Patient Results..."
        pBar.Value = 16
        lblStatus.Refresh
        myConnection.Execute "delete from tbPatientValues"
        lblStatus.Caption = "Clean-up Instrument Results..."
        pBar.Value = 17
        lblStatus.Refresh
        myConnection.Execute "delete from tbLabInstrumentResults"
        lblStatus.Caption = "Clean-up Instrument Result Values..."
        pBar.Value = 17
        lblStatus.Refresh
        myConnection.Execute "delete from tbLabInstrumentRValues"
        lblStatus.Caption = "Clean-up Exams..."
        pBar.Value = 18
        lblStatus.Refresh
        myConnection.Execute "delete from tbCoLabExam"
        lblStatus.Caption = "Clean-up Class..."
        pBar.Value = 19
        lblStatus.Refresh
        myConnection.Execute "delete from tbLabClassifications"
        lblStatus.Caption = "Clean-up Class Values..."
        pBar.Value = 20
        lblStatus.Refresh
        myConnection.Execute "delete from tbLabClassExams"
        lblStatus.Caption = "Clean-up Profiles..."
        pBar.Value = 21
        lblStatus.Refresh
        myConnection.Execute "delete from tbLabProfiles"
        lblStatus.Caption = "Clean-up Profile Exams..."
        pBar.Value = 22
        lblStatus.Refresh
        myConnection.Execute "delete from tbLabProfileExams"
        lblStatus.Caption = "Clean-up Hospital Information..."
        pBar.Value = 23
        lblStatus.Refresh
        myConnection.Execute "delete from tbHospitalInfo"
        lblStatus.Caption = "Clean-up Pathologist..."
        pBar.Value = 24
        lblStatus.Refresh
        myConnection.Execute "delete from tbLabPathologists"
        lblStatus.Caption = "Clean-up Physician..."
        pBar.Value = 25
        lblStatus.Refresh
        myConnection.Execute "delete from tbPhysician"
        lblStatus.Caption = "Clean-up Med.Tech..."
        pBar.Value = 26
        lblStatus.Refresh
        myConnection.Execute "delete from tbMedtech"
        lblStatus.Caption = "Clean-up Complete..."
        pBar.Value = 27
        lblStatus.Refresh
        
        pBar.Value = 28
        lblStatus.Refresh
        myConnection.Execute "insert into tbpatientmaster select * from tempMaster"
        lblStatus.Caption = "Uploading Patient Results..."
        pBar.Value = 29
        lblStatus.Refresh
        myConnection.Execute "insert into tbPatientValues select * from tempValues"
        lblStatus.Caption = "Uploading Instrument Results..."
        pBar.Value = 30
        lblStatus.Refresh
        myConnection.Execute "insert into tbLabInstrumentResults Select InstrumentID," & _
            " ResultDate, Accessionnum, PatientName, Status, ProfileExam, RefNum, CassNum," & _
            " RackNum, Printed, NewPatientName, AccessionnumOrg, IDnum, BirthDate, Doctor," & _
            " Room, Age, Sex, Ward, Specimen from tempResults where  instr(1,refnum,'N') > 0"
        
        myConnection.Execute "insert into tbLabInstrumentResults Select InstrumentID," & _
            " ResultDate, Accessionnum, PatientName, Status, ProfileExam, RefNum, CassNum," & _
            " RackNum, Printed, NewPatientName, AccessionnumOrg, IDnum, BirthDate, Doctor," & _
            " Room, Age, Sex, Ward, Specimen from tempResults where sortnum = " & txtHeaderCode.Text & ""

        GenerateInsertResultID
        lblStatus.Caption = "Uploading Exams..."
        pBar.Value = 31
        lblStatus.Refresh
        myConnection.Execute "insert into tbCoLabExam select * from tempExam"
        lblStatus.Caption = "Uploading Class..."
        pBar.Value = 32
        lblStatus.Refresh
        myConnection.Execute "insert into tbLabClassifications select * from tempClass"
        lblStatus.Caption = "Uploading Class Values..."
        pBar.Value = 33
        lblStatus.Refresh
        myConnection.Execute "insert into tbLabClassExams select * from tempClassVal"
        lblStatus.Caption = "Uploading Profiles..."
        pBar.Value = 34
        lblStatus.Refresh
        If findCol("tempprofile", "INSTRUMENTID") Then
            myConnection.Execute "insert into tbLabProfiles select * from tempProfile"
        Else
            myConnection.Execute "insert into tbLabProfiles select *,'" & InstrumentID & "' as InstrumentID from tempProfile"
        End If
        lblStatus.Caption = "Uploading Profile Exams..."
        pBar.Value = 35
        lblStatus.Refresh
        If findCol("tempProfileVal", "INSTRUMENTID") Then
            myConnection.Execute "insert into tbLabProfileExams select * from tempProfileVal"
        Else
            myConnection.Execute "insert into tbLabProfileExams select *,'" & InstrumentID & "' as InstrumentID from tempProfileVal"
        End If
        lblStatus.Caption = "Uploading Hospital Information..."
        pBar.Value = 36
        lblStatus.Refresh
        myConnection.Execute "insert into tbHospitalInfo select * from tempHospital"
        lblStatus.Caption = "Uploading Pathologist..."
        pBar.Value = 37
        lblStatus.Refresh
        myConnection.Execute "insert into tbLabPathologists select * from tempPathologists"
        lblStatus.Caption = "Uploading Physician..."
        pBar.Value = 38
        lblStatus.Refresh
        myConnection.Execute "insert into tbPhysician select * from tempPhysician"
        lblStatus.Caption = "Uploading Med.Tech..."
        pBar.Value = 39
        lblStatus.Refresh
        myConnection.Execute "insert into tbMedtech select * from tempMedtech"
        lblStatus.Caption = "Uploading Complete..."
        pBar.Value = 40
        lblStatus.Refresh

        myConnection.Execute "drop table tempMaster"
        lblStatus.Caption = "Removing Temporary Tables"
        pBar.Value = 41
        lblStatus.Refresh
        myConnection.Execute "drop table tempValues"
        lblStatus.Caption = "Removing Temporary Tables"
        pBar.Value = 42
        lblStatus.Refresh
        myConnection.Execute "drop table tempResults"
        lblStatus.Caption = "Removing Temporary Tables"
        pBar.Value = 43
        lblStatus.Refresh
        myConnection.Execute "drop table tempExam"
        lblStatus.Caption = "Removing Temporary Tables"
        pBar.Value = 44
        lblStatus.Refresh
        myConnection.Execute "drop table tempClass"
        lblStatus.Caption = "Removing Temporary Tables"
        pBar.Value = 45
        lblStatus.Refresh
        myConnection.Execute "drop table tempClassVal"
        lblStatus.Caption = "Removing Temporary Tables"
        pBar.Value = 46
        lblStatus.Refresh
        myConnection.Execute "drop table tempProfile"
        lblStatus.Caption = "Removing Temporary Tables"
        pBar.Value = 47
        lblStatus.Refresh
        myConnection.Execute "drop table tempProfileVal"
        lblStatus.Caption = "Removing Temporary Tables"
        pBar.Value = 48
        lblStatus.Refresh
        myConnection.Execute "drop table tempHospital"
        lblStatus.Caption = "Done"
        pBar.Value = 49
        lblStatus.Refresh
        myConnection.Execute "drop table tempPathologists"
        lblStatus.Caption = "Removing Temporary Tables"
        pBar.Value = 50
        lblStatus.Refresh
        myConnection.Execute "drop table tempPhysician"
        lblStatus.Caption = "Removing Temporary Tables"
        pBar.Value = 51
        lblStatus.Refresh
        myConnection.Execute "drop table tempMedtech"
        lblStatus.Caption = "Done"
        pBar.Value = 52
        lblStatus.Refresh
        myConnection.CommitTrans
        Me.MousePointer = 0
    End If
End If
End Sub

Private Function findCol(mtable As String, SearchField As String) As Boolean
Dim rstTemp As ADODB.Recordset
Dim ctr As Integer
findCol = False
Set rstTemp = myConnection.Execute("Select * from " & mtable)
For ctr = 0 To rstTemp.Fields.Count - 1
    If Format(rstTemp.Fields(ctr).Name, ">") = Format(SearchField, ">") Then
        findCol = True
        Exit For
    End If
Next ctr
End Function

Private Sub GenerateInsertResultID()
Dim rstResults As ADODB.Recordset
Dim myResultID As String
Dim ctr As Long
Set rstResults = myConnection.Execute("Select Distinct InstrumentID, Resultdate, PatientName, AccessionnumOrg, SortNum, ExamName, Result, Unit, ResultStr, RefNum from tempResults")
If rstResults.RecordCount > 0 Then
    rstResults.MoveLast
    rstResults.MoveFirst
    pbResultID.Visible = True
    pbResultID.Min = 1
    pbResultID.Max = rstResults.RecordCount + 1
    ctr = 1
    lblText(2).Visible = True
    DoEvents
    Me.Refresh
    myConnection.BeginTrans
    Do While Not rstResults.EOF
        lblText(2).Refresh
        myResultID = GenerateResultID(rstResults.Fields("InstrumentID").Value & rstResults.Fields("ResultDate").Value & rstResults.Fields("PatientName").Value & rstResults.Fields("AccessionNumOrg").Value)
        'myResultID = GenerateResultID(rstResults.Fields("ResultDate").Value & Format(Now, "MMddyyhhmmss"))
'        myResultID = Blowfish.EncryptString(rstResults.Fields("ResultDate").Value, myKey, True)
        If rstResults.Fields("SortNum").Value = Val(txtHeaderCode.Text) Then
            myConnection.Execute "Update tbLabInstrumentResults set ResultID = '" & myResultID & "' " & _
                "where InstrumentID = '" & _
                rstResults.Fields("InstrumentID").Value & "' and PatientName = '" & Replace(rstResults.Fields("PatientName").Value, "'", "''") & "' and AccessionNumOrg = '" & _
                rstResults.Fields("AccessionNumOrg").Value & "' and Resultdate = #" & rstResults.Fields("ResultDate").Value & "#"
        ElseIf InStr(1, rstResults.Fields("RefNum").Value, "N") > 0 Then
            myConnection.Execute "Update tbLabInstrumentResults set ResultID = '" & myResultID & "' " & _
                "where refnum = '" & rstResults.Fields("RefNum").Value & "'"
        End If
        If Not IsNull(rstResults.Fields("Result").Value) Then
            myConnection.Execute "insert into tbLabInstrumentRValues (ResultID, ExamName, Result, Unit, ResultStr, SortNum)" & _
            "select '" & myResultID & "','" & rstResults.Fields("ExamName").Value & _
                "'," & rstResults.Fields("Result").Value & ",'" & rstResults.Fields("Unit").Value & _
                "',''," & IIf(IsNull(rstResults.Fields("SortNum").Value), 0, rstResults.Fields("SortNum").Value) & ";"
        ElseIf InStr(1, rstResults.Fields("RefNum").Value, "N") > 0 Then
            myConnection.Execute "insert into tbLabInstrumentRValues (ResultID, SortNum)" & _
            "select '" & myResultID & "'," & txtHeaderCode.Text & ";"
        Else
            myConnection.Execute "insert into tbLabInstrumentRValues (ResultID, SortNum)" & _
            "select '" & myResultID & "','" & IIf(IsNull(rstResults.Fields("SortNum").Value), 0, rstResults.Fields("SortNum").Value) & "';"
        End If
        rstResults.MoveNext
        pbResultID.Value = ctr
        lblText(2).Caption = "Copying " & ctr & "/" & pbResultID.Max
        lblText(2).Refresh
        ctr = ctr + 1
    Loop
    myConnection.CommitTrans
    pbResultID.Visible = False
    lblText(2).Visible = False
End If
End Sub

Private Sub Text2_Change()

End Sub

Private Sub txtMajor_Change()
    If Len(Trim(txtMajor.Text)) > 0 Then cmdStart.Enabled = True
End Sub
