VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmDbMan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Manager"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   4125
   StartUpPosition =   1  'CenterOwner
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1035
      Left            =   5580
      TabIndex        =   19
      Top             =   420
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   1826
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      Caption         =   "Update Database Password"
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
      Height          =   1875
      Index           =   3
      Left            =   60
      TabIndex        =   14
      Top             =   4860
      Width           =   3975
      Begin VB.TextBox txtOldPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1620
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   390
         Width           =   2115
      End
      Begin VB.TextBox txtNewPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1620
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   720
         Width           =   2115
      End
      Begin VB.TextBox txtConfirmPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1620
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1050
         Width           =   2115
      End
      Begin VB.CommandButton cmdChangePassword 
         Caption         =   "Change Passsword"
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
         Left            =   1980
         TabIndex        =   8
         Top             =   1440
         Width           =   1785
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password:"
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
         Left            =   510
         TabIndex        =   17
         Top             =   450
         Width           =   1080
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Password:"
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
         Left            =   390
         TabIndex        =   16
         Top             =   780
         Width           =   1185
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password:"
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
         Left            =   150
         TabIndex        =   15
         Top             =   1080
         Width           =   1395
      End
   End
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5850
      TabIndex        =   10
      Top             =   2940
      Width           =   3975
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database Manager"
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
         TabIndex        =   11
         Top             =   150
         Width           =   1845
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   4785
      Index           =   0
      Left            =   60
      TabIndex        =   12
      Top             =   30
      Width           =   3975
      Begin VB.CommandButton cmdGenerateResultID 
         Caption         =   "Migrate Instrument Results"
         Height          =   405
         Left            =   180
         TabIndex        =   24
         Top             =   4230
         Width           =   3615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy Data from Old Database..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   180
         TabIndex        =   23
         Top             =   3450
         Width           =   3615
      End
      Begin VB.CommandButton cmdRestoreDB 
         Caption         =   "Restore Database from Backup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   180
         TabIndex        =   3
         Top             =   2310
         Width           =   3615
      End
      Begin VB.CommandButton cmdClearBuffer 
         Caption         =   "Clear Host Spooler"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   180
         TabIndex        =   0
         Top             =   630
         Width           =   3615
      End
      Begin VB.CommandButton cmdArchive 
         Caption         =   "Archive Samples"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   180
         TabIndex        =   4
         Top             =   2880
         Width           =   3615
      End
      Begin VB.CommandButton cmdFixLink 
         Caption         =   "Link Database"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   180
         TabIndex        =   1
         Top             =   1200
         Width           =   3615
      End
      Begin VB.CommandButton cmdCompactDB 
         Caption         =   "Compact and Backup Database"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   180
         TabIndex        =   2
         Top             =   1740
         Width           =   3615
      End
      Begin MSComctlLib.ProgressBar pbrStatus 
         Height          =   135
         Left            =   210
         TabIndex        =   20
         Top             =   4290
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Left            =   180
         TabIndex        =   21
         Top             =   4020
         Width           =   3555
      End
      Begin VB.Label lblText 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   750
         TabIndex        =   22
         Top             =   210
         Width           =   2475
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   795
      Index           =   1
      Left            =   5610
      TabIndex        =   13
      Top             =   1770
      Width           =   3975
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   765
      Index           =   2
      Left            =   60
      TabIndex        =   18
      Top             =   6750
      Width           =   3975
      Begin VB.CommandButton cmdDBpatch 
         Caption         =   "Repatch DB"
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
         Left            =   90
         TabIndex        =   25
         Top             =   180
         Width           =   1575
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
         Height          =   465
         Left            =   1680
         TabIndex        =   9
         Top             =   180
         Width           =   2235
      End
   End
End
Attribute VB_Name = "frmDbMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cForms As New ClassForms
Private CompactOnly As Boolean

Private Sub cmdArchive_Click()
Dim myArray() As String
Dim tempStr As String
Me.MousePointer = 11
tempStr = cForms.showArchive
If Len(Trim(tempStr)) > 0 Then
    myArray = Split(tempStr, "|")
    BackupDatabase Format(Date, "yyyymmdd") & "ARC", False
    myConnection.Execute "delete * from tbLabInstrumentRValues where resultid in (select distinct resultid from tbLabInstrumentResults where resultdate >= #" & Format(myArray(0), "mm/dd/yyyy 00:00") & "# and  resultdate <= #" & Format(myArray(1), "mm/dd/yyyy 23:59") & "#)"
    myConnection.Execute "delete * from tbLabInstrumentResults where resultdate >= #" & Format(myArray(0), "mm/dd/yyyy 00:00") & "# and  resultdate <= #" & Format(myArray(1), "mm/dd/yyyy 23:59") & "#"
    myConnection.Execute "delete * from tbpatientvalues where refnum in (select refnum from tbpatientmaster where resultdate >= #" & Format(myArray(0), "mm/dd/yyyy 00:00") & "# and  resultdate <= #" & Format(myArray(1), "mm/dd/yyyy 23:59") & "# and refnum <> 'TESTPREVIEW')"
    myConnection.Execute "delete * from tbPatientMaster where resultdate >= #" & Format(myArray(0), "mm/dd/yyyy 00:00") & "# and  resultdate <= #" & Format(myArray(1), "mm/dd/yyyy 23:59") & "# and refnum <> 'TESTPREVIEW'"
    MsgBox "Archiving Complete...", vbOKOnly + vbInformation, "Archive"
End If
Me.MousePointer = 0
End Sub

Private Sub cmdChangePassword_Click()
Dim cnn1 As New ADODB.Connection
Dim dbLoc As String
On Error GoTo errorFound
dbLoc = Replace(Format(DBDatabase, ">"), "LABORATORY.MDB", "")
DoEvents
cmdChangePassword.Enabled = False
showStatus True, 1, 12, "Closing Connections..."
CloseConnection

showStatus True, 1, 12, "Updating Password..."
cnn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbLoc & "\Laboratory.mdb" & ";User Id=" & "" & ";Password=" & "" & ";" & "Jet OLEDB:Database Password=" & txtOldPassword.Text & ";"
cnn1.ConnectionTimeout = 30
cnn1.CursorLocation = adUseClient
cnn1.Mode = adModeShareExclusive
cnn1.Open
cnn1.Execute "alter database password " & IIf(Len(Trim(txtNewPassword.Text)) > 0, txtNewPassword.Text, "``") & " " & IIf(Len(Trim(txtOldPassword.Text)) > 0, txtOldPassword.Text, "``")
cnn1.Close

showStatus True, 2, 12, "Updating Password..."
cnn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbLoc & "\InstrumentResults.mdb" & ";User Id=" & "" & ";Password=" & "" & ";" & "Jet OLEDB:Database Password=" & txtOldPassword.Text & ";"
cnn1.ConnectionTimeout = 30
cnn1.CursorLocation = adUseClient
cnn1.Mode = adModeShareExclusive
cnn1.Open
cnn1.Execute "alter database password " & IIf(Len(Trim(txtNewPassword.Text)) > 0, txtNewPassword.Text, "``") & " " & IIf(Len(Trim(txtOldPassword.Text)) > 0, txtOldPassword.Text, "``")
cnn1.Close

showStatus True, 3, 12, "Updating Password..."
cnn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbLoc & "\LabBuffer.mdb" & ";User Id=" & "" & ";Password=" & "" & ";" & "Jet OLEDB:Database Password=" & txtOldPassword.Text & ";"
cnn1.ConnectionTimeout = 30
cnn1.CursorLocation = adUseClient
cnn1.Mode = adModeShareExclusive
cnn1.Open
cnn1.Execute "alter database password " & IIf(Len(Trim(txtNewPassword.Text)) > 0, txtNewPassword.Text, "``") & " " & IIf(Len(Trim(txtOldPassword.Text)) > 0, txtOldPassword.Text, "``")
cnn1.Close

showStatus True, 4, 12, "Updating Password..."
cnn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbLoc & "\LabPatientData.mdb" & ";User Id=" & "" & ";Password=" & "" & ";" & "Jet OLEDB:Database Password=" & txtOldPassword.Text & ";"
cnn1.ConnectionTimeout = 30
cnn1.CursorLocation = adUseClient
cnn1.Mode = adModeShareExclusive
cnn1.Open
cnn1.Execute "alter database password " & IIf(Len(Trim(txtNewPassword.Text)) > 0, txtNewPassword.Text, "``") & " " & IIf(Len(Trim(txtOldPassword.Text)) > 0, txtOldPassword.Text, "``")
cnn1.Close

showStatus True, 5, 12, "Updating Password..."
cnn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbLoc & "\LabPatientValues.mdb" & ";User Id=" & "" & ";Password=" & "" & ";" & "Jet OLEDB:Database Password=" & txtOldPassword.Text & ";"
cnn1.ConnectionTimeout = 30
cnn1.CursorLocation = adUseClient
cnn1.Mode = adModeShareExclusive
cnn1.Open
cnn1.Execute "alter database password " & IIf(Len(Trim(txtNewPassword.Text)) > 0, txtNewPassword.Text, "``") & " " & IIf(Len(Trim(txtOldPassword.Text)) > 0, txtOldPassword.Text, "``")
cnn1.Close

showStatus True, 6, 12, "Updating Password..."
cnn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbLoc & "\InstrumentRValues.mdb" & ";User Id=" & "" & ";Password=" & "" & ";" & "Jet OLEDB:Database Password=" & txtOldPassword.Text & ";"
cnn1.ConnectionTimeout = 30
cnn1.CursorLocation = adUseClient
cnn1.Mode = adModeShareExclusive
cnn1.Open
cnn1.Execute "alter database password " & IIf(Len(Trim(txtNewPassword.Text)) > 0, txtNewPassword.Text, "``") & " " & IIf(Len(Trim(txtOldPassword.Text)) > 0, txtOldPassword.Text, "``")
cnn1.Close

showStatus True, 6, 12, "Updating Password..."
cnn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbLoc & "\LabRequestBuffer.mdb" & ";User Id=" & "" & ";Password=" & "" & ";" & "Jet OLEDB:Database Password=" & txtOldPassword.Text & ";"
cnn1.ConnectionTimeout = 30
cnn1.CursorLocation = adUseClient
cnn1.Mode = adModeShareExclusive
cnn1.Open
cnn1.Execute "alter database password " & IIf(Len(Trim(txtNewPassword.Text)) > 0, txtNewPassword.Text, "``") & " " & IIf(Len(Trim(txtOldPassword.Text)) > 0, txtOldPassword.Text, "``")
cnn1.Close

showStatus True, 7, 12, "Encrypting..."
If Len(Trim(txtNewPassword.Text)) > 0 Then SaveSetting "LAB-i", "LOCAL", "DATABASEPASSWORD", Blowfish.EncryptString(txtNewPassword.Text, myKey, True) Else SaveSetting "LAB-i", "LOCAL", "DATABASEPASSWORD", ""

showStatus True, 8, 12, "Retrieving Settings..."
GetDBSettings
showStatus True, 9, 12, "Opening Connections..."
OpenConnection
showStatus True, 10, 12, "Fixing Attatchments..."
FixLink DBDatabase, False
showStatus True, 11, 12, "Initializing Host Forms..."
setHostForms
showStatus True, 12, 12, "Initializing..."
txtOldPassword.Text = ""
txtNewPassword.Text = ""
txtConfirmPassword.Text = ""
showStatus True, 12, 12, "Done."
MsgBox "Password has been successfully changed!", vbInformation + vbOKOnly, App.ProductName
showStatus False
errorFound:
If Err.Number <> 0 Then
    MsgBox Err.Number & ": " & Err.Description
End If
End Sub

Private Sub cmdClearBuffer_Click()
On Error Resume Next
Me.MousePointer = 11
fraBody(0).Enabled = False
If MsgBox("Are you sure you want to clear the instrument database buffer?", vbCritical + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
    myConnection.RollbackTrans
    myConnection.Execute "delete from LabUnprocessed where instrumentid = '" & InstrumentID & "'"
    MsgBox "Instrument Database Buffer has been cleared.", vbOKOnly + vbInformation, App.ProductName
End If
fraBody(0).Enabled = True
Me.MousePointer = 0
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCompactDB_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 1 Then CompactOnly = True Else CompactOnly = False
End Sub

Private Sub cmdCompactDB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mFilename As String
    If Shift = 1 Then
        Me.MousePointer = 11
        CloseConnection
        DAO.DBEngine.CompactDatabase Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "Laboratory.mdb", Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "tLaboratory.mdb", dbLangGeneral, , ";pwd=" & DBPassword
        DAO.DBEngine.CompactDatabase Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "InstrumentResults.mdb", Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "tInstrumentResults.mdb", dbLangGeneral, , ";pwd=" & DBPassword
        DAO.DBEngine.CompactDatabase Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "LabBuffer.mdb", Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "tLabBuffer.mdb", dbLangGeneral, , ";pwd=" & DBPassword
        DAO.DBEngine.CompactDatabase Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "LabPatientData.mdb", Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "tLabPatientData.mdb", dbLangGeneral, , ";pwd=" & DBPassword
        DAO.DBEngine.CompactDatabase Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "LabPatientValues.mdb", Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "tLabPatientValues.mdb", dbLangGeneral, , ";pwd=" & DBPassword
        DAO.DBEngine.CompactDatabase Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "InstrumentRValues.mdb", Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "tInstrumentRValues.mdb", dbLangGeneral, , ";pwd=" & DBPassword
        fsoFiles.CopyFile Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "tLaboratory.mdb", Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "Laboratory.mdb", True
        fsoFiles.CopyFile Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "tInstrumentResults.mdb", Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "InstrumentResults.mdb", True
        fsoFiles.CopyFile Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "tLabBuffer.mdb", Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "LabBuffer.mdb", True
        fsoFiles.CopyFile Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "tLabPatientData.mdb", Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "LabPatientData.mdb", True
        fsoFiles.CopyFile Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "tLabPatientValues.mdb", Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "LabPatientValues.mdb", True
        fsoFiles.CopyFile Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "tInstrumentRValues.mdb", Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "InstrumentRValues.mdb", True
        fsoFiles.DeleteFile Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "tLaboratory.mdb", True
        fsoFiles.DeleteFile Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "tInstrumentResults.mdb", True
        fsoFiles.DeleteFile Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "tLabBuffer.mdb", True
        fsoFiles.DeleteFile Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "tLabPatientData.mdb", True
        fsoFiles.DeleteFile Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "tLabPatientValues.mdb", True
        fsoFiles.DeleteFile Replace(lblText(0).ToolTipText, "Laboratory.mdb", "") & "tInstrumentRValues.mdb", True
        OpenConnection
        Me.MousePointer = 0
        MsgBox "Compact and Repair Complete. Press Ok to continue.", vbInformation + vbOKOnly, App.ProductName
    Else
        Me.MousePointer = 11
        mFilename = cForms.showInputBox("Filename", , , Format(Date, "yyyymmdd")) ' Format(Date, "yyyymmdd")
        If Len(Trim(mFilename)) > 0 Then
            BackupDatabase mFilename, True
        End If
        Me.MousePointer = 0
    End If
End Sub

Private Sub cmdCopy_Click()
frmDataBaseCopy.Show 1
End Sub

Private Sub cmdDBpatch_Click()
    SaveSetting "Lab-I", "Settings", "UpdatedTable", "0"
    MsgBox "Database Patch has been requested. Please restart the application.", vbOKOnly + vbInformation, "DB Patch"
End Sub

Private Sub cmdFixLink_Click()
    DoEvents
    Me.MousePointer = 11
    showStatus True, 100, 100, "Attaching Database..."
    fraBody(0).Enabled = False
    FixLink DBDatabase, False
    lblStatus.Caption = "Done."
    MsgBox "Fixed Database Link.", vbOKOnly + vbInformation, App.ProductName
    fraBody(0).Enabled = True
    showStatus False
    Me.MousePointer = 0
End Sub

Private Sub showStatus(willShow As Boolean, Optional curCnt As Integer = 100, Optional totCnt As Integer = 100, Optional tStatus As String = "")
If willShow Then
    lblStatus.Visible = True
    pbrStatus.Visible = True
    pbrStatus.Max = totCnt
    pbrStatus.Value = curCnt
    lblStatus.Caption = tStatus
Else
    lblStatus.Visible = False
    pbrStatus.Visible = False
End If
End Sub

Private Sub cmdGenerateResultID_Click()
Dim rstResults As ADODB.Recordset
Dim tbPatientValues_Crosstab As String
Dim myResultID As String
Dim ctr As Long

tbPatientValues_Crosstab = "PARAMETERS mRefNum Text ( 255 );  " & _
"TRANSFORM First(tbPatientValues.Result) AS FirstOfResult " & _
"SELECT tbPatientValues.RefNum " & _
"FROM tbPatientValues " & _
"WHERE tbPatientValues.RefNum = mRefNum " & _
"GROUP BY tbPatientValues.RefNum " & _
"PIVOT tbPatientValues.DisplayName;"

If MsgBox("This will generate ResultID on TBLABINSTRUMENTRESULTS on Older Database file. Do you wish to continue?", vbYesNo + vbQuestion + vbDefaultButton2, "Generate Result ID") = vbYes Then
    If MsgBox("Do you want to take responsibility of any damage done on database?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        cmdGenerateResultID.Visible = False
        pbrStatus.Visible = True
        lblStatus.Visible = True
        If MsgBox("Do you want to clear Result ID?", vbYesNo + vbQuestion + vbDefaultButton2, "Reset Result ID") = vbYes Then
            Me.MousePointer = 11
            DoEvents
            lblStatus.Caption = "Updating RESULTID..."
            myConnection.Execute "update tbLabInstrumentResults set resultid = '' where refnum = resultid"
        End If
        Me.MousePointer = 11
        Set rstResults = myConnection.Execute("Select Distinct InstrumentID, Resultdate, PatientName, AccessionnumOrg from tbLabInstrumentResults")
        If rstResults.RecordCount > 0 Then
            rstResults.MoveFirst
            myConnection.BeginTrans
            pbrStatus.Max = rstResults.RecordCount
            ctr = 1
            Do While Not rstResults.EOF
                DoEvents
                pbrStatus.Value = rstResults.AbsolutePosition
                lblStatus.Caption = "Processing:" & ctr & "/" & rstResults.RecordCount
                myResultID = GenerateResultID(rstResults.Fields("InstrumentID").Value & rstResults.Fields("ResultDate").Value & rstResults.Fields("PatientName").Value & rstResults.Fields("AccessionNumOrg").Value)
                myConnection.Execute "Update tbLabInstrumentResults set ResultID = '" & myResultID & "' " & _
                    "where InstrumentID = '" & _
                    rstResults.Fields("InstrumentID").Value & "' and PatientName = '" & Replace(rstResults.Fields("PatientName").Value, "'", "''") & "' and AccessionNumOrg = '" & _
                    rstResults.Fields("AccessionNumOrg").Value & "' and Resultdate = #" & rstResults.Fields("ResultDate").Value & "#"
                rstResults.MoveNext
                If ctr Mod 10000 = 0 Then
                    myConnection.CommitTrans
                    
                    lblStatus.Caption = "Copying data to tbLabInstrumentRValues..."
                    myConnection.BeginTrans
                    myConnection.Execute "insert into tbLabInstrumentRValues (ResultID, ExamName, Result, Unit,  ResultStr, SortNum) select ResultID, ExamName, Result, Unit,  ResultStr, SortNum from tbLabInstrumentResults where len(trim(sortnum)) > 0 and len(trim(resultid)) > 0"
                    myConnection.CommitTrans
                    
                    lblStatus.Caption = "Deleting Resultvalues from tbLabInstrumentResults..."
                    myConnection.BeginTrans
                    myConnection.Execute "delete from tbLabInstrumentResults where sortnum > 1 and len(trim(resultid)) > 0"
                    myConnection.CommitTrans
                    
                    myConnection.BeginTrans
                End If
                ctr = ctr + 1
            Loop
            myConnection.CommitTrans
        End If
        lblStatus.Caption = "Copying data to tbLabInstrumentRValues..."
        myConnection.BeginTrans
        myConnection.Execute "insert into tbLabInstrumentRValues (ResultID, ExamName, Result, Unit,  ResultStr, SortNum) select ResultID, ExamName, Result, Unit,  ResultStr, SortNum from tbLabInstrumentResults where len(trim(sortnum)) > 0 and len(trim(resultid)) > 0"
        myConnection.CommitTrans
        
        lblStatus.Caption = "Deleting Resultvalues from tbLabInstrumentResults..."
        myConnection.BeginTrans
        myConnection.Execute "delete from tbLabInstrumentResults where sortnum > 1 and len(trim(resultid)) > 0"
        myConnection.CommitTrans
        
        myConnection.Execute "DROP Procedure tbPatientValues_Crosstab"
        myConnection.Execute "CREATE Procedure tbPatientValues_Crosstab as " & tbPatientValues_Crosstab
        Me.MousePointer = 0
        
        If MsgBox("Migration complete. Do you want to Backup/Compact Database?", vbYesNo + vbQuestion + vbDefaultButton2, "Migration Complete") = vbYes Then
            BackupDatabase Format("yyyymmdd" & "_MG"), True
        End If
        
        cmdGenerateResultID.Visible = True
        pbrStatus.Visible = False
        lblStatus.Visible = False
    End If
End If
Me.MousePointer = 0
End Sub

Private Sub cmdRestoreDB_Click()
Dim dirStr As String
Dim strCurDir As String
Dim strSrcDir As String
Dim tempStr As String
    dirStr = showBackupList
    If Len(Trim(dirStr)) > 0 Then
        If MsgBox("Are you sure you want to restore this backup?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
            strSrcDir = Replace(Format(DBDatabase, ">"), "LABORATORY.MDB", "")
            strCurDir = strSrcDir
            strSrcDir = strSrcDir & "Backup\" & dirStr & "\"

            If Right(dirStr, 3) <> "RES" Then
                   BackupDatabase Format(Date, "yyyymmdd") & "RES", True
            End If
            
            CloseConnection
            fsoFiles.CopyFile strSrcDir & "Laboratory.mdb", strCurDir & "Laboratory.mdb", True
            fsoFiles.CopyFile strSrcDir & "InstrumentResults.mdb", strCurDir & "InstrumentResults.mdb", True
            fsoFiles.CopyFile strSrcDir & "LabBuffer.mdb", strCurDir & "LabBuffer.mdb", True
            fsoFiles.CopyFile strSrcDir & "LabPatientData.mdb", strCurDir & "LabPatientData.mdb", True
            fsoFiles.CopyFile strSrcDir & "LabPatientValues.mdb", strCurDir & "LabPatientValues.mdb", True
            fsoFiles.CopyFile strSrcDir & "InstrumentRValues.mdb", strCurDir & "InstrumentRValues.mdb", True
            GetDBSettings
            OpenConnection
            FixLink DBDatabase, False
            setHostForms
            MsgBox "Restoration Successful. Press Ok to continue.", vbInformation + vbOKOnly, App.ProductName
        End If
    End If
End Sub

Private Sub Form_Load()
Dim rstVersion As ADODB.Recordset
setFormColors Me
CompactOnly = False
On Error GoTo errFound
pbrStatus.Visible = False
lblStatus.Visible = False
txtOldPassword.Text = ""
txtNewPassword.Text = ""
txtConfirmPassword.Text = ""
cmdChangePassword.Enabled = False
Set rstVersion = myConnection.Execute("Select * from tbVersion")
lblText(0).Caption = "DB Version " & rstVersion.Fields("Major").Value & "." & rstVersion.Fields("Minor").Value
lblText(0).ToolTipText = DBDatabase
errFound:
    If Err.Number = -2147217865 Then
        lblText(0).Caption = "DB Version 1.0"
    ElseIf Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description
    End If
End Sub

Private Sub lblText_DblClick(Index As Integer)
    Shell "explorer " + Replace(Format(DBDatabase, ">"), "LABORATORY.MDB", ""), vbNormalFocus
End Sub

Private Sub MSHFlexGrid1_DblClick()
Dim ctr As Integer
With MSHFlexGrid1
    For ctr = 1 To .Rows - 1
        myConnection.Execute "Insert into tblabclassexams (InstrumentID, ClassID, ExamID, RangeUp, RangeDown, RangeUnit, Display, Active, RangeUpSI, RangeDownSI, RangeUnitSI, DisplaySI )" & _
            "Select '" & .TextMatrix(ctr, 1) & "','" & .TextMatrix(ctr, 2) & "','" & .TextMatrix(ctr, 3) & "'," & .TextMatrix(ctr, 7) & "," & .TextMatrix(ctr, 8) & ",'" & .TextMatrix(ctr, 4) & "','" & .TextMatrix(ctr, 5) & "',TRUE," & .TextMatrix(ctr, 9) & "," & .TextMatrix(ctr, 10) & ",'" & .TextMatrix(ctr, 12) & "','" & .TextMatrix(ctr, 11) & "'"
    Next ctr
    MsgBox "OK"
End With
End Sub

Private Sub txtConfirmPassword_Change()
    If txtNewPassword.Text = txtConfirmPassword.Text And txtOldPassword.Text = DBPassword Then
        cmdChangePassword.Enabled = True
    Else
        cmdChangePassword.Enabled = False
    End If
End Sub

Private Sub txtNewPassword_Change()
    If txtNewPassword.Text = txtConfirmPassword.Text And txtOldPassword.Text = DBPassword Then
        cmdChangePassword.Enabled = True
    Else
        cmdChangePassword.Enabled = False
    End If
End Sub

Private Sub txtOldPassword_Change()
    If txtNewPassword.Text = txtConfirmPassword.Text And txtOldPassword.Text = DBPassword Then
        cmdChangePassword.Enabled = True
    Else
        cmdChangePassword.Enabled = False
    End If
End Sub

