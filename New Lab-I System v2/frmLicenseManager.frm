VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmLicenseManager 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Host Interface License"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkDirectPort 
      Caption         =   "Connect using proxy port"
      Height          =   315
      Left            =   30
      TabIndex        =   12
      Top             =   1020
      Width           =   2475
   End
   Begin VB.Timer tmrTimeout 
      Interval        =   1000
      Left            =   480
      Top             =   2760
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9060
      TabIndex        =   2
      Top             =   570
      Width           =   2025
   End
   Begin VB.TextBox txtSampleNum 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1590
      TabIndex        =   0
      Top             =   540
      Width           =   7425
   End
   Begin MSWinsockLib.Winsock wsCodeServer 
      Left            =   30
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "License"
      Height          =   3315
      Left            =   30
      TabIndex        =   4
      Top             =   1410
      Width           =   11025
      Begin VB.ListBox lstLicenses 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   3090
         TabIndex        =   9
         Top             =   1380
         Width           =   4905
      End
      Begin VB.CommandButton cmdGetLicense 
         Caption         =   "Get License"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4440
         TabIndex        =   6
         Top             =   2730
         Width           =   2175
      End
      Begin VB.Label lblInstrument 
         Alignment       =   2  'Center
         Caption         =   "Instrument"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   60
         TabIndex        =   11
         Top             =   930
         Width           =   10935
      End
      Begin VB.Label lblRegisteredTo 
         Alignment       =   2  'Center
         Caption         =   "Registered Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   60
         TabIndex        =   10
         Top             =   510
         Width           =   10935
      End
      Begin VB.Label lblNolicense 
         Alignment       =   2  'Center
         Caption         =   "Please connect to the internet then click Get License"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   8
         Top             =   990
         Width           =   10935
      End
      Begin VB.Label lblNolicense 
         Alignment       =   2  'Center
         Caption         =   "No License Found"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   540
         Width           =   10935
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   4770
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19632
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Host Interface Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   60
      TabIndex        =   3
      Top             =   90
      Width           =   10995
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      Caption         =   "HOST INTERFACE ID"
      Height          =   210
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   1470
   End
End
Attribute VB_Name = "frmLicenseManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnConnected As Boolean
Private CancelCount As Integer
Private mHWID1 As String
Private mHWID2 As String
Private mHWID3 As String
Private mHOSTRAM As String
Private mHOSTSA As String
Private mHOSTEXE As String
Private blnGetLicense As Boolean

Private Sub chkDirectPort_Click()
    SaveSetting "LAB-i", "Registration", "ProxyPort", chkDirectPort.Value
End Sub

Private Sub cmdGetLicense_Click()
    cmdGetLicense.Enabled = False
    cmdRegister_Click
End Sub

Private Sub cmdRegister_Click()
If wsCodeServer.State = 0 Then
    OpenRegistrationServer mHWID1, mHWID2, mHWID3, mHOSTRAM, mHOSTSA, mHOSTEXE
Else
    MsgBox "Connection is still busy...", vbOKOnly + vbInformation, "Busy"
End If
End Sub

Private Sub Form_Load()
    blnGetLicense = False
    mHWID1 = Replace(GetCpuID, "-", "")
    mHWID2 = Replace(GetSerialNumber(Left(App.Path, 3)), "-", "")
    mHWID3 = Replace(SystemSerialNumber, "-", "")
    mHOSTRAM = GetTotalRAM
    mHOSTSA = GetSAVersion
    mHOSTEXE = App.Major & "." & App.Minor & "." & App.Revision
    LoadDetails
    chkDirectPort.Value = Val(GetSetting("LAB-i", "Registration", "ProxyPort", "1"))
End Sub

Private Function LoadDetails() As String
Dim StrValues As String
Dim ArrValues() As String
Dim ArrLicense() As String
Dim iFileNo As Integer
Dim mDirectory As String
Dim tempstr As String
Dim newStr As String
Dim ctr As Integer
iFileNo = FreeFile
StrValues = GetSetting("LAB-i", "Registration", "HOSTUID", "")
If Trim(StrValues) = "" Then
    lblStatus.Caption = "Unregistered Host Computer"
    cmdRegister.Enabled = True
    lblNolicense(0).Visible = False
    lblNolicense(1).Visible = False
    cmdGetLicense.Visible = False
    lstLicenses.Visible = False
    lblInstrument.Visible = False
    lblRegisteredTo.Visible = False
    lblInstrument.Visible = False
    lblRegisteredTo.Visible = False
Else
    cmdRegister.Enabled = False
    txtSampleNum.Text = StrValues
    lblStatus.Caption = "Registered"
    lblNolicense(0).Visible = True
    lblNolicense(1).Visible = True
    cmdGetLicense.Visible = True
    lblInstrument.Visible = False
    lblRegisteredTo.Visible = False
    mDirectory = Replace(Format(DBDatabase, ">"), "LABORATORY.MDB", "")
    If Len(Dir(mDirectory & "\" & Environ("USERNAME") & "@" & Environ("USERDOMAIN"), vbDirectory)) = 0 Then
        fsoFiles.Createfolder mDirectory & "\" & Environ("USERNAME") & "@" & Environ("USERDOMAIN")
    End If
    mDirectory = mDirectory & "\" & Environ("USERNAME") & "@" & Environ("USERDOMAIN")
    If Len(Dir(mDirectory & "\Laboratory.dat")) > 0 Then
        Open mDirectory & "\Laboratory.dat" For Input As #iFileNo
        Do While Not EOF(iFileNo)
          Input #iFileNo, tempstr
        Loop
        Close #iFileNo
        newStr = tempstr
        newStr = Blowfish.DecryptString(newStr, myKey, True)
        newStr = Mid(newStr, 101)
        If InStr(1, newStr, "|") > 0 Then
            ArrValues = Split(newStr, "|")
            lblInstrument.Visible = True
            lblRegisteredTo.Visible = True
            lblRegisteredTo.Caption = ArrValues(5)
            lblInstrument.Caption = ArrValues(6)
            lblNolicense(0).Visible = False
            lblNolicense(1).Visible = False
            ArrLicense = Split(ArrValues(7), "<SP>")
            lstLicenses.Clear
            For ctr = 0 To UBound(ArrLicense)
                lstLicenses.AddItem Replace(ArrLicense(ctr), "^", "-")
            Next ctr
            lstLicenses.Visible = True
        End If
    End If
End If
End Function

Private Function GetLicences(mHWUID As String) As String
Dim Requestheader As String
Dim PayloadData As String
Dim mHost As String
wsCodeServer.Protocol = sckTCPProtocol
wsCodeServer.RemoteHost = "mdmpi.com.ph"
If Val(GetSetting("LAB-i", "Registration", "ProxyPort", "1")) = 1 Then
    wsCodeServer.RemotePort = 80
Else
    wsCodeServer.RemotePort = 9998
End If
wsCodeServer.RemoteHost = "localhost"
'wsCodeServer.RemotePort = 9998
wsCodeServer.Connect
mHost = wsCodeServer.RemoteHost & ":" & wsCodeServer.RemotePort
While Not blnConnected And CancelCount < 30
    DoEvents
    Me.MousePointer = 11
    StatusBar1.Panels(1).Text = "Connecting to server(" & CancelCount & ")..."
Wend
If blnConnected Then
    StatusBar1.Panels(1).Text = "Connected. Waiting for response..."
Else
    StatusBar1.Panels(1).Text = "Unable to establish connection..."
End If
If Not blnConnected And CancelCount >= 30 Then
    wsCodeServer.Close
    wsCodeServer.LocalPort = 0
    CancelCount = 0
    wsCodeServer.Protocol = sckTCPProtocol
    wsCodeServer.RemoteHost = "mdmpi.com.ph"
    wsCodeServer.RemotePort = 9998
    wsCodeServer.Connect
    mHost = wsCodeServer.RemoteHost & ":" & wsCodeServer.RemotePort
    While Not blnConnected And CancelCount < 30
        DoEvents
        Me.MousePointer = 11
        StatusBar1.Panels(1).Text = "Retrying via proxy(" & CancelCount & ")..."
    Wend
    If blnConnected Then
        StatusBar1.Panels(1).Text = "Connected via proxy. Waiting for response..."
    Else
        StatusBar1.Panels(1).Text = "Unable to connect via proxy..."
    End If
End If
PayloadData = "HWUID=" & mHWUID
PayloadData = Blowfish.EncryptString(PayloadData, myKey, True)
Requestheader = "POST /api_gethostlicenses HTTP/1.0" & vbCrLf
Requestheader = Requestheader & "Host: " & mHost & vbCrLf
Requestheader = Requestheader & "Accept-Language: en-US,en;q=0.5" & vbCrLf
Requestheader = Requestheader & "Content-Type: text/plain" & vbCrLf
Requestheader = Requestheader & "Content-Length: " & Len(PayloadData) & vbCrLf
Requestheader = Requestheader & "DNT: 1" & vbCrLf
Requestheader = Requestheader & "Connection: keep -alive" & vbCrLf
Requestheader = Requestheader & "Upgrade-Insecure-Requests: 1" & vbCrLf

CancelCount = 1
If blnConnected And CancelCount < 30 Then
    wsCodeServer.SendData Requestheader & vbCrLf & PayloadData
Else
    StatusBar1.Panels(1).Text = "Cannot find server..."
    If wsCodeServer.State <> 0 Then
        wsCodeServer.Close
        wsCodeServer.LocalPort = 0
    End If
End If
Me.MousePointer = 0
End Function

Private Sub Form_Unload(Cancel As Integer)
    RegisterOK = True
End Sub

Private Sub OpenRegistrationServer(mHWID1 As String, mHWID2 As String, mHWID3 As String, mHOSTRAM As String, mHOSTSA As String, mHOSTEXE As String)
Dim Requestheader As String
Dim PayloadData As String
Dim mHost As String
wsCodeServer.Protocol = sckTCPProtocol
wsCodeServer.RemoteHost = "mdmpi.com.ph"
If Val(GetSetting("LAB-i", "Registration", "ProxyPort", "1")) = 1 Then
    wsCodeServer.RemotePort = 80
Else
    wsCodeServer.RemotePort = 9998
End If
'wsCodeServer.RemoteHost = "localhost"
'wsCodeServer.RemotePort = 9998
wsCodeServer.Connect
mHost = wsCodeServer.RemoteHost & ":" & wsCodeServer.RemotePort
While Not blnConnected And CancelCount < 30
    DoEvents
    Me.MousePointer = 11
    StatusBar1.Panels(1).Text = "Connecting to server(" & CancelCount & ")..."
Wend
If blnConnected Then
    StatusBar1.Panels(1).Text = "Connected. Waiting for response..."
Else
    StatusBar1.Panels(1).Text = "Unable to establish connection..."
End If
'If Not blnConnected And CancelCount >= 30 Then
'    wsCodeServer.Close
'    CancelCount = 0
'    wsCodeServer.Protocol = sckTCPProtocol
'    wsCodeServer.RemoteHost = "mdmpi.com.ph"
'    wsCodeServer.RemotePort = 9998
'    wsCodeServer.Connect
'    mHost = wsCodeServer.RemoteHost & ":" & wsCodeServer.RemotePort
'    While Not blnConnected And CancelCount < 30
'        DoEvents
'        Me.MousePointer = 11
'        StatusBar1.Panels(1).Text = "Retrying via proxy(" & CancelCount & ")..."
'    Wend
'    If blnConnected Then
'        StatusBar1.Panels(1).Text = "Connected via proxy. Waiting for response..."
'    Else
'        StatusBar1.Panels(1).Text = "Unable to connect via proxy..."
'    End If
'End If
PayloadData = "HWID1=" & mHWID1 & "&HWID2=" & mHWID2 & "&HWID3=" & mHWID3 & "&HOSTRAM=" & mHOSTRAM & "&HOSTSA=" & mHOSTSA & "&HOSTEXE=" & mHOSTEXE & ""
PayloadData = Blowfish.EncryptString(PayloadData, myKey, True)
Requestheader = "POST /api_registerhost HTTP/1.0" & vbCrLf
Requestheader = Requestheader & "Host: " & mHost & vbCrLf
Requestheader = Requestheader & "Accept-Language: en-US,en;q=0.5" & vbCrLf
Requestheader = Requestheader & "Content-Type: text/plain" & vbCrLf
Requestheader = Requestheader & "Content-Length: " & Len(PayloadData) & vbCrLf
Requestheader = Requestheader & "DNT: 1" & vbCrLf
Requestheader = Requestheader & "Connection: keep -alive" & vbCrLf
Requestheader = Requestheader & "Upgrade-Insecure-Requests: 1" & vbCrLf

CancelCount = 1
If blnConnected And CancelCount < 30 Then
    wsCodeServer.SendData Requestheader & vbCrLf & PayloadData
Else
    StatusBar1.Panels(1).Text = "Cannot find server..."
    If wsCodeServer.State <> 0 Then
        wsCodeServer.Close
        wsCodeServer.LocalPort = 0
    End If
End If
Me.MousePointer = 0
End Sub

Private Sub tmrTimeOut_Timer()
    CancelCount = CancelCount + 1
    Select Case wsCodeServer.State
        Case 1
            blnConnected = True
            If blnGetLicense Then
                blnGetLicense = False
                GetLicences txtSampleNum.Text
            End If
        Case 9
            blnConnected = False
            wsCodeServer.Close
            wsCodeServer.LocalPort = 0
            wsCodeServer.Connect
    End Select
End Sub

Private Sub wsCodeServer_Close()
    wsCodeServer.Close
    wsCodeServer.LocalPort = 0
    blnConnected = False
    If blnGetLicense Then
        blnGetLicense = False
        GetLicences txtSampleNum.Text
    End If
    StatusBar1.Panels(1).Text = "Done"
End Sub

Private Sub wsCodeServer_Connect()
    blnConnected = True
End Sub

Private Sub wsCodeServer_DataArrival(ByVal bytesTotal As Long)
    Dim strResponse As String
    wsCodeServer.GetData strResponse, vbString, bytesTotal
    ProcessResponse strResponse
End Sub

Private Sub ProcessResponse(mMess As String)
Dim PayloadData As String
Dim arrTemp() As String
Dim tempstr As String
Dim ArrLicense() As String
Dim strDate As String
Dim ctr As Integer
PayloadData = Mid(mMess, InStr(1, mMess, vbCrLf & vbCrLf) + 4)
PayloadData = Blowfish.DecryptString(PayloadData, myKey, True)
Debug.Print PayloadData
If Len(PayloadData) > 0 And InStr(1, PayloadData, "|") > 0 Then
    arrTemp = Split(PayloadData, "|")
    Select Case arrTemp(0)
        Case "HOSTREG"
            If mHWID1 = arrTemp(2) And mHWID2 = arrTemp(3) And mHWID3 = arrTemp(4) Then
                SaveSetting "LAB-i", "Registration", "HOSTUID", Replace(Replace(arrTemp(1), "{", ""), "}", "")
                SaveSetting "LAB-i", "Registration", "HOSTWID1", arrTemp(2)
                SaveSetting "LAB-i", "Registration", "HOSTWID2", arrTemp(3)
                SaveSetting "LAB-i", "Registration", "HOSTWID3", arrTemp(4)
                tempstr = arrTemp(1) & "|" & arrTemp(2) & "|" & arrTemp(3) & "|" & arrTemp(4) & "|||"
                txtSampleNum.Text = GetSetting("LAB-i", "Registration", "HOSTUID", "")
                blnGetLicense = True
'                GetLicences txtSampleNum.Text
                cmdGetLicense.Enabled = False
                cmdRegister.Enabled = False
                lblNolicense(0).Visible = True
                lblNolicense(1).Visible = True
                cmdGetLicense.Visible = True
                lblStatus.Caption = "Registered"
                DeletePayload
            End If
        Case "HOSTLICENSE"
            If arrTemp(7) = "^" Then
                MsgBox "No License found. Please contact Marsman Drysdale Medical Products, Inc. as soon as possible."
                DeletePayload
                lblNolicense(0).Visible = True
                lblNolicense(1).Visible = True
                lblRegisteredTo.Visible = False
                lblInstrument.Visible = False
                lstLicenses.Clear
            Else
                SavePayload PayloadData
                strDate = ""
                LoadDetails
            End If
            cmdGetLicense.Enabled = True
    End Select
End If
End Sub

Private Sub SavePayload(mPayloadData As String)
Dim mDirectory As String
Dim mFile As String
Dim iFileNo As Integer
Dim NewPayloadData As String
NewPayloadData = Blowfish.EncryptString(String(100, "-") & mPayloadData, myKey, True)
iFileNo = FreeFile
mDirectory = Replace(Format(DBDatabase, ">"), "LABORATORY.MDB", "")
If Len(Dir(mDirectory & Environ("USERNAME") & "@" & Environ("USERDOMAIN"), vbDirectory)) = 0 Then
    fsoFiles.Createfolder mDirectory & "\" & Environ("USERNAME") & "@" & Environ("USERDOMAIN")
End If
mDirectory = mDirectory & Environ("USERNAME") & "@" & Environ("USERDOMAIN")
Open mDirectory & "\Laboratory.dat" For Output As #iFileNo
Print #iFileNo, NewPayloadData
Close #iFileNo
End Sub

Private Sub DeletePayload()
Dim mDirectory As String
Dim mFile As String
mDirectory = Replace(Format(DBDatabase, ">"), "LABORATORY.MDB", "")
mDirectory = mDirectory & "\" & Environ("USERNAME") & "@" & Environ("USERDOMAIN")
If Len(Dir(mDirectory & "Laboratory.dat")) > 0 Then
    fsoFiles.DeleteFile mDirectory & "Laboratory.dat"
End If
End Sub
