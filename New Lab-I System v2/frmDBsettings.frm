VERSION 5.00
Begin VB.Form frmDBsettings 
   BackColor       =   &H006B4B01&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Settings"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   3840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboServerType 
      Height          =   330
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   405
      Left            =   2610
      TabIndex        =   12
      Top             =   3450
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   405
      Left            =   1440
      TabIndex        =   11
      Top             =   3450
      Width           =   1095
   End
   Begin VB.Frame fraBody 
      BackColor       =   &H006B4B01&
      Caption         =   "CONNECTION SETTINGS"
      ForeColor       =   &H00FFFFFF&
      Height          =   1995
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   1290
      Width           =   3615
      Begin VB.TextBox txtUsername 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         TabIndex        =   6
         Top             =   330
         Width           =   2385
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1110
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   720
         Width           =   2385
      End
      Begin VB.TextBox txtDatabase 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         TabIndex        =   4
         Top             =   1110
         Width           =   2385
      End
      Begin VB.TextBox txtSchema 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         TabIndex        =   3
         Top             =   1500
         Width           =   2385
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   360
         TabIndex        =   10
         Top             =   1140
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   300
         TabIndex        =   9
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   270
         TabIndex        =   8
         Top             =   390
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Schema:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   450
         TabIndex        =   7
         Top             =   1530
         Width           =   630
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      Caption         =   "Server Name"
      ForeColor       =   &H80000008&
      Height          =   705
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   510
      Width           =   3675
      Begin VB.TextBox txtServerName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   210
         TabIndex        =   1
         Top             =   270
         Width           =   3255
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SERVER TYPE:"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   120
      TabIndex        =   14
      Top             =   150
      Width           =   1095
   End
End
Attribute VB_Name = "frmDBsettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fillcboServerType()
With cboServerType
    .AddItem "LOCAL DATABASE"
    .AddItem "MSSQL SERVER"
    .AddItem "ORACLE"
    .AddItem "MYSQL SERVER"
End With
End Sub

Private Sub cboServerType_Click()
Dim tempStr As String
Select Case cboServerType.Text
    Case "LOCAL DATABASE"
        fraBody(1).Caption = "LOCAL DATABASE"
        txtUsername.Text = ""
        txtPassword.Text = ""
        txtDatabase.Text = ""
        txtServerName.Text = ""
        txtSchema.Text = ""
        txtUsername.Enabled = False
        txtPassword.Enabled = False
        txtDatabase.Enabled = False
        fraBody(0).Enabled = False
        fraBody(1).Enabled = False
        txtSchema.Enabled = False
    Case "MSSQL SERVER"
        fraBody(1).Caption = "SQL SERVER CONNECTION SETTINGS"
        tempStr = GetSetting("lab-i", "MSSQL", "UserName", "")
        If Len(Trim(tempStr)) > 0 Then
            txtUsername = Blowfish.DecryptString(tempStr, myKey, True)
        Else
            txtUsername = ""
        End If
        tempStr = GetSetting("lab-i", "MSSQL", "Password", "")
        If Len(Trim(tempStr)) > 0 Then
            txtPassword = Blowfish.DecryptString(tempStr, myKey, True)
        Else
            txtPassword = ""
        End If
        tempStr = GetSetting("lab-i", "MSSQL", "Database", "")
        If Len(Trim(tempStr)) > 0 Then
            txtDatabase = Blowfish.DecryptString(tempStr, myKey, True)
        Else
            txtDatabase = ""
        End If
        tempStr = GetSetting("lab-i", "MSSQL", "Server", Blowfish.EncryptString("(LOCAL)", myKey, True))
        If Len(Trim(tempStr)) > 0 Then
            txtServerName.Text = Blowfish.DecryptString(tempStr, myKey, True)
        Else
            txtServerName.Text = "(LOCAL)"
        End If

        txtUsername.Enabled = True
        txtPassword.Enabled = True
        txtDatabase.Enabled = True
        txtSchema.Enabled = False
        txtSchema.Text = ""
        fraBody(0).Enabled = True
        fraBody(1).Enabled = True
    Case "MYSQL SERVER"
        fraBody(1).Caption = "MYSQL SERVER CONNECTION SETTINGS"
        tempStr = GetSetting("lab-i", "MYSQL", "UserName", "")
        If Len(Trim(tempStr)) > 0 Then
            txtUsername = Blowfish.DecryptString(tempStr, myKey, True)
        Else
            txtUsername = ""
        End If
        tempStr = GetSetting("lab-i", "MYSQL", "Password", "")
        If Len(Trim(tempStr)) > 0 Then
            txtPassword = Blowfish.DecryptString(tempStr, myKey, True)
        Else
            txtPassword = ""
        End If
        tempStr = GetSetting("lab-i", "MYSQL", "Database", "")
        If Len(Trim(tempStr)) > 0 Then
            txtDatabase = Blowfish.DecryptString(tempStr, myKey, True)
        Else
            txtDatabase = ""
        End If
        tempStr = GetSetting("lab-i", "MYSQL", "Server", Blowfish.EncryptString("(LOCAL)", myKey, True))
        If Len(Trim(tempStr)) > 0 Then
            txtServerName.Text = Blowfish.DecryptString(tempStr, myKey, True)
        Else
            txtServerName.Text = "(LOCAL)"
        End If

        txtUsername.Enabled = True
        txtPassword.Enabled = True
        txtDatabase.Enabled = True
        txtSchema.Enabled = False
        txtSchema.Text = ""
        fraBody(0).Enabled = True
        fraBody(1).Enabled = True
    Case "ORACLE"
        fraBody(1).Caption = "ORACLE CONNECTION SETTINGS"
        tempStr = GetSetting("lab-i", "ORACLE", "UserName", "")
        If Len(Trim(tempStr)) > 0 Then
            txtUsername = Blowfish.DecryptString(tempStr, myKey, True)
        Else
            txtUsername = ""
        End If
        tempStr = GetSetting("lab-i", "ORACLE", "Password", "")
        If Len(Trim(tempStr)) > 0 Then
            txtPassword = Blowfish.DecryptString(tempStr, myKey, True)
        Else
            txtPassword = ""
        End If
        tempStr = GetSetting("lab-i", "ORACLE", "Database", "")
        If Len(Trim(tempStr)) > 0 Then
            txtDatabase = Blowfish.DecryptString(tempStr, myKey, True)
        Else
            txtDatabase = ""
        End If
        tempStr = GetSetting("lab-i", "ORACLE", "Schema", "")
        If Len(Trim(tempStr)) > 0 Then
            txtSchema = Blowfish.DecryptString(tempStr, myKey, True)
        Else
            txtSchema = ""
        End If
        tempStr = GetSetting("lab-i", "ORACLE", "Server", Blowfish.EncryptString("(LOCAL)", myKey, True))
        If Len(Trim(tempStr)) > 0 Then
            txtServerName.Text = Blowfish.DecryptString(tempStr, myKey, True)
        Else
            txtServerName.Text = "(LOCAL)"
        End If
        
        txtUsername.Enabled = True
        txtPassword.Enabled = True
        txtDatabase.Enabled = True
        fraBody(0).Enabled = True
        txtSchema.Enabled = True
        fraBody(1).Enabled = True
End Select
End Sub

Private Sub cboServerType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
Select Case cboServerType.Text
    Case "LOCAL DATABASE"
        SaveSetting "lab-i", "DBSettings", "ServerType", Blowfish.EncryptString("LOCAL DATABASE", myKey, True)
    Case "MSSQL SERVER"
        SaveSetting "lab-i", "DBSettings", "ServerType", Blowfish.EncryptString("MSSQL SERVER", myKey, True)
        If Len(Trim(txtUsername.Text)) > 0 Then SaveSetting "lab-i", "MSSQL", "UserName", Blowfish.EncryptString(txtUsername.Text, myKey, True) Else SaveSetting "lab-i", "MSSQL", "UserName", ""
        If Len(Trim(txtPassword.Text)) > 0 Then SaveSetting "lab-i", "MSSQL", "Password", Blowfish.EncryptString(txtPassword.Text, myKey, True) Else SaveSetting "lab-i", "MSSQL", "Password", ""
        If Len(Trim(txtDatabase.Text)) > 0 Then SaveSetting "lab-i", "MSSQL", "Database", Blowfish.EncryptString(txtDatabase.Text, myKey, True) Else SaveSetting "lab-i", "MSSQL", "Database", ""
        If Len(Trim(txtServerName.Text)) > 0 Then SaveSetting "lab-i", "MSSQL", "Server", Blowfish.EncryptString(txtServerName.Text, myKey, True) Else SaveSetting "lab-i", "MSSQL", "Server", Blowfish.EncryptString("(LOCAL)", myKey, True)
    Case "MYSQL SERVER"
        SaveSetting "lab-i", "DBSettings", "ServerType", Blowfish.EncryptString("MYSQL SERVER", myKey, True)
        If Len(Trim(txtUsername.Text)) > 0 Then SaveSetting "lab-i", "MYSQL", "UserName", Blowfish.EncryptString(txtUsername.Text, myKey, True) Else SaveSetting "lab-i", "MYSQL", "UserName", ""
        If Len(Trim(txtPassword.Text)) > 0 Then SaveSetting "lab-i", "MYSQL", "Password", Blowfish.EncryptString(txtPassword.Text, myKey, True) Else SaveSetting "lab-i", "MYSQL", "Password", ""
        If Len(Trim(txtDatabase.Text)) > 0 Then SaveSetting "lab-i", "MYSQL", "Database", Blowfish.EncryptString(txtDatabase.Text, myKey, True) Else SaveSetting "lab-i", "MYSQL", "Database", ""
        If Len(Trim(txtServerName.Text)) > 0 Then SaveSetting "lab-i", "MYSQL", "Server", Blowfish.EncryptString(txtServerName.Text, myKey, True) Else SaveSetting "lab-i", "MYSQL", "Server", Blowfish.EncryptString("(LOCAL)", myKey, True)
    Case "ORACLE"
        SaveSetting "lab-i", "DBSettings", "ServerType", Blowfish.EncryptString("ORACLE", myKey, True)
        If Len(Trim(txtUsername.Text)) > 0 Then SaveSetting "lab-i", "ORACLE", "UserName", Blowfish.EncryptString(txtUsername.Text, myKey, True) Else SaveSetting "lab-i", "ORACLE", "UserName", ""
        If Len(Trim(txtPassword.Text)) > 0 Then SaveSetting "lab-i", "ORACLE", "Password", Blowfish.EncryptString(txtPassword.Text, myKey, True) Else SaveSetting "lab-i", "ORACLE", "Password", ""
        If Len(Trim(txtDatabase.Text)) > 0 Then SaveSetting "lab-i", "ORACLE", "Database", Blowfish.EncryptString(txtDatabase.Text, myKey, True) Else SaveSetting "lab-i", "ORACLE", "Database", ""
        If Len(Trim(txtSchema.Text)) > 0 Then SaveSetting "lab-i", "ORACLE", "Schema", Blowfish.EncryptString(txtSchema.Text, myKey, True) Else SaveSetting "lab-i", "ORACLE", "Schema", ""
        If Len(Trim(txtServerName.Text)) > 0 Then SaveSetting "lab-i", "ORACLE", "Server", Blowfish.EncryptString(txtServerName.Text, myKey, True) Else SaveSetting "lab-i", "ORACLE", "Server", Blowfish.EncryptString("(LOCAL)", myKey, True)
End Select
GetDBSettings
MsgBox "Settings has been succesfully saved!", vbOKOnly + vbInformation, App.ProductName
End Sub

Private Sub Form_Load()
    fillcboServerType
    setCBOServer DBServerType
    setFormColors Me
End Sub

Private Sub setCBOServer(findString As String)
Dim ctr As Integer
Dim Count As Integer
Dim found As Boolean
found = False
With cboServerType
    For ctr = 0 To .ListCount - 1
        .ListIndex = ctr
        If .Text = findString Then
            found = True
            Count = .ListIndex
            Exit For
        End If
    Next ctr
    If found Then .ListIndex = Count Else .ListIndex = 0
End With
End Sub

Private Sub txtDatabase_GotFocus()
    Highlight
End Sub

Private Sub txtDatabase_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtDatabase_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPassword_GotFocus()
    Highlight
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtSchema_GotFocus()
    Highlight
End Sub

Private Sub txtSchema_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtSchema_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtServerName_GotFocus()
    Highlight
End Sub

Private Sub txtServerName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtServerName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtUsername_GotFocus()
    Highlight
End Sub

Private Sub txtUsername_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub
