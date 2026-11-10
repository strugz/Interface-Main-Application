VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmRegistration 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Laboratory Instrument Interface Registration"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8055
   ClipControls    =   0   'False
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
   ScaleHeight     =   5130
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTempCode 
      Alignment       =   2  'Center
      Height          =   795
      Left            =   270
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "frmRegistration.frx":0000
      Top             =   3270
      Width           =   7395
   End
   Begin VB.CommandButton cmdNewCode 
      Caption         =   "Enter New Code"
      Height          =   555
      Left            =   1620
      TabIndex        =   18
      Top             =   4140
      Width           =   1935
   End
   Begin VB.Timer tmrTimeout 
      Interval        =   1000
      Left            =   540
      Top             =   4170
   End
   Begin MSWinsockLib.Winsock wsCodeServer 
      Left            =   60
      Top             =   4170
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtHospitalName 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   270
      MaxLength       =   250
      TabIndex        =   5
      Top             =   1020
      Width           =   7395
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   4755
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14155
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
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Height          =   555
      Left            =   3630
      TabIndex        =   6
      Top             =   4140
      Width           =   2805
   End
   Begin VB.TextBox txtCode5 
      Height          =   315
      Left            =   5070
      MaxLength       =   12
      TabIndex        =   4
      Text            =   "A52EE0FD56BC"
      Top             =   360
      Width           =   1425
   End
   Begin VB.TextBox txtCode4 
      Height          =   315
      Left            =   4320
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "B044"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtCode3 
      Height          =   315
      Left            =   3570
      MaxLength       =   4
      TabIndex        =   2
      Text            =   "4AA8"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtCode2 
      Height          =   315
      Left            =   2790
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "EB56"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtCode1 
      Height          =   315
      Left            =   1590
      MaxLength       =   8
      TabIndex        =   0
      Text            =   "4550156D"
      Top             =   360
      Width           =   1065
   End
   Begin VB.Label lblTempCode 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Temporary Code:"
      Height          =   210
      Left            =   3390
      TabIndex        =   20
      Top             =   3060
      Width           =   1275
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Host Interface Code"
      Height          =   225
      Left            =   1620
      TabIndex        =   17
      Top             =   90
      Width           =   4845
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Hospital Name"
      Height          =   225
      Left            =   1620
      TabIndex        =   16
      Top             =   780
      Width           =   4845
   End
   Begin VB.Label lblExpiration 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1620
      TabIndex        =   15
      Top             =   2580
      Width           =   4815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Expiration Date:"
      Height          =   210
      Left            =   1620
      TabIndex        =   14
      Top             =   2340
      Width           =   4815
   End
   Begin VB.Label lblRegname 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1620
      TabIndex        =   12
      Top             =   1830
      Width           =   4815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Registered To:"
      Height          =   210
      Left            =   1620
      TabIndex        =   11
      Top             =   1590
      Width           =   4830
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   210
      Left            =   4980
      TabIndex        =   10
      Top             =   390
      Width           =   60
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   210
      Left            =   4230
      TabIndex        =   9
      Top             =   390
      Width           =   60
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   210
      Left            =   3450
      TabIndex        =   8
      Top             =   390
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   210
      Left            =   2700
      TabIndex        =   7
      Top             =   390
      Width           =   60
   End
End
Attribute VB_Name = "frmRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnConnected As Boolean
Private CancelCount As Integer
Private blnOpenTemp As Boolean

Private Sub cmdNewCode_Click()
If MsgBox("Are you sure you want to enter new code? This will clear the current code.", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
    EnControls 1, 1, 1, 0
    txtCode1.Text = ""
    txtCode2.Text = ""
    txtCode3.Text = ""
    txtCode4.Text = ""
    txtCode5.Text = ""
    lblRegname.Caption = ""
    txtHospitalName.Text = ""
    lblExpiration.Caption = "-"
    SaveSetting "Lab-I", "Registration", "Code1", txtCode1.Text
    SaveSetting "Lab-I", "Registration", "Code2", txtCode2.Text
    SaveSetting "Lab-I", "Registration", "Code3", txtCode3.Text
    SaveSetting "Lab-I", "Registration", "Code4", txtCode4.Text
    SaveSetting "Lab-I", "Registration", "Code5", txtCode5.Text
    SaveSetting "Lab-I", "Registration", "Registered Name", lblRegname.Caption
    SaveSetting "Lab-I", "Registration", "Hospital Name", txtHospitalName.Text
    SaveSetting "Lab-I", "Registration", "Expiration", "-"
    fsoFiles.DeleteFile Replace(Format(DBDatabase, ">"), "LABORATORY.MDB", "") & "\Laboratory.dat"
End If
End Sub

Private Sub cmdRegister_Click()
Dim mCode As String
Dim mDate As String
Dim mHDDserial As String
Dim mBoardSerial As String
Dim mRegName As String
StatusBar1.Panels(1).Text = ""
CancelCount = 1
mDate = Format(Now, "yyyymmddHHmmss")
mCode = txtCode1.Text & "-" & txtCode2.Text & "-" & txtCode3.Text & "-" & txtCode4.Text & "-" & txtCode5.Text
mHDDserial = GetSerialNumber(Left(App.Path, 3))
mBoardSerial = SystemSerialNumber
mHDDserial = Replace(mHDDserial, "-", "")
mBoardSerial = Replace(mBoardSerial, "-", "")
If lblRegname.Caption = "-" Then
    mRegName = ""
Else
    mRegName = lblRegname.Caption
End If
OpenRegistrationServer mCode, mRegName, txtHospitalName.Text, mHDDserial, mBoardSerial, mDate, blnOpenTemp
blnOpenTemp = False
End Sub

Private Sub cmdRegister_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift = 3 Then
    blnOpenTemp = True
End If
End Sub

Private Sub Form_Load()
    txtCode1.Text = GetSetting("Lab-I", "Registration", "Code1", "")
    txtCode2.Text = GetSetting("Lab-I", "Registration", "Code2", "")
    txtCode3.Text = GetSetting("Lab-I", "Registration", "Code3", "")
    txtCode4.Text = GetSetting("Lab-I", "Registration", "Code4", "")
    txtCode5.Text = GetSetting("Lab-I", "Registration", "Code5", "")
    lblRegname.Caption = GetSetting("Lab-I", "Registration", "Registered Name", "-")
    txtHospitalName.Text = GetSetting("Lab-I", "Registration", "Hospital Name", "")
    lblExpiration.Caption = Format(GetSetting("Lab-I", "Registration", "Expiration", "-"), "MMMM dd, yyyy")
    txtTempCode.Text = GetSetting("Lab-I", "Registration", "TempCode", "")
    blnOpenTemp = False
    If Len(Trim(txtTempCode.Text)) = 0 Then
         lblTempCode.Visible = False
         txtTempCode.Visible = False
    End If
    If IsRegistered Then
        RegisterOK = True
        EnControls 0, 0, 0, 1
    Else
        cmdRegister.Enabled = ValidateCode
        If IsDate(lblExpiration.Caption) Then
            If DateDiff("d", Now, lblExpiration.Caption) <= 0 Then
                If lblRegname.Caption = "-" Then
                    StatusBar1.Panels(1).Text = "Please register the system. Registration Code and Internet connection is required."
                    EnControls 1, 1, 1, 0
                Else
                    StatusBar1.Panels(1).Text = "Laboratory Instrument Interface has reached its expiration date."
                    EnControls 0, 0, 1, 0
                End If
            Else
                StatusBar1.Panels(1).Text = "System is having some difficulties validating activation. Please re-register the system."
                cmdRegister.Caption = "Retrieve Registration"
                EnControls 0, 0, 1, 0
            End If
        Else
            If lblExpiration.Caption = "-" Then
                StatusBar1.Panels(1).Text = "Please register the system. Registration Code and Internet connection is required."
                EnControls 1, 1, 1, 0
            Else
                StatusBar1.Panels(1).Text = "System is having some difficulties validating activation. Please re-register the system."
                EnControls 0, 0, 1, 0
            End If
        End If
    End If
End Sub

Private Function ValidateCode() As Boolean
Dim tempBool As Boolean
tempBool = False
If Len(txtCode1.Text) = 8 And Len(txtCode2.Text) = 4 And Len(txtCode3.Text) = 4 And Len(txtCode4.Text) = 4 And Len(txtCode5.Text) = 12 And Len(txtHospitalName.Text) > 0 Then
    If ISHEX(txtCode1.Text) And ISHEX(txtCode2.Text) And ISHEX(txtCode3.Text) And ISHEX(txtCode4.Text) And ISHEX(txtCode5.Text) Then
        tempBool = True
    End If
End If
ValidateCode = tempBool
End Function

Private Function ISHEX(mess As String) As Boolean
Dim HexValue As String
Dim tempBool As Boolean
Dim ctr As Integer
tempBool = True
HexValue = "0123456789ABCDEF"
For ctr = 1 To Len(mess)
    If InStr(1, HexValue, Mid(mess, ctr, 1)) = 0 Then
        tempBool = False
        Exit For
    End If
Next ctr
ISHEX = tempBool
End Function

Private Sub Form_Unload(Cancel As Integer)
'RegisterOK = True
End Sub

Private Sub tmrTimeOut_Timer()
    CancelCount = CancelCount + 1
    Select Case wsCodeServer.State
        Case 1
            blnConnected = True
        Case 9
            blnConnected = False
            wsCodeServer.Close
            wsCodeServer.Connect
    End Select
End Sub

Private Sub txtCode1_Change()
    cmdRegister.Enabled = ValidateCode
    If Len(txtCode1.Text) = txtCode1.MaxLength Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtCode1_GotFocus()
    Highlight
End Sub

Private Sub txtCode1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCode2_Change()
    cmdRegister.Enabled = ValidateCode
    If Len(txtCode2.Text) = txtCode2.MaxLength Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtCode2_GotFocus()
Highlight
End Sub

Private Sub txtCode2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCode3_Change()
    cmdRegister.Enabled = ValidateCode
    If Len(txtCode3.Text) = txtCode3.MaxLength Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtCode3_GotFocus()
Highlight
End Sub

Private Sub txtCode3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCode4_Change()
    cmdRegister.Enabled = ValidateCode
    If Len(txtCode4.Text) = txtCode4.MaxLength Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtCode4_GotFocus()
Highlight
End Sub

Private Sub txtCode4_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCode5_Change()
    cmdRegister.Enabled = ValidateCode
    If Len(txtCode5.Text) = txtCode5.MaxLength Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtCode5_GotFocus()
Highlight
End Sub

Private Sub txtCode5_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtHospitalName_Change()
    cmdRegister.Enabled = ValidateCode
End Sub

Private Sub OpenRegistrationServer(mRegCode As String, mRegisteredName As String, mHospName As String, mHW1 As String, mHW2 As String, mReqDate As String, mOpenTemp As Boolean)
Dim Requestheader As String
Dim PayloadData As String
Dim mHost As String
EnControls 0, 0, 0, 0
wsCodeServer.Protocol = sckTCPProtocol
wsCodeServer.RemoteHost = "mdmpi.com.ph"
wsCodeServer.RemotePort = 80
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
PayloadData = "INTERFACECODE=" & mRegCode & "&REGISTEREDNAME=" & mRegisteredName & "&HOSPNAME=" & mHospName & "&HWID1=" & mHW1 & "&HWID2=" & mHW2 & "&DATEREQUESTED=" & mReqDate & ""
PayloadData = Blowfish.EncryptString(PayloadData, myKey, True)
Requestheader = "POST /api_interfaceregistration HTTP/1.0" & vbCrLf
Requestheader = Requestheader & "Host: " & mHost & vbCrLf
Requestheader = Requestheader & "Accept-Language: en-US,en;q=0.5" & vbCrLf
Requestheader = Requestheader & "Content-Type: text/plain" & vbCrLf
Requestheader = Requestheader & "Content-Length: " & Len(PayloadData) & vbCrLf
Requestheader = Requestheader & "DNT: 1" & vbCrLf
Requestheader = Requestheader & "Connection: keep -alive" & vbCrLf
Requestheader = Requestheader & "Upgrade-Insecure-Requests: 1" & vbCrLf

If CancelCount >= 30 And mOpenTemp Then
    lblTempCode.Visible = True
    txtTempCode.Visible = True
End If
CancelCount = 1
If blnConnected And CancelCount < 30 Then
    wsCodeServer.SendData Requestheader & vbCrLf & PayloadData
Else
    EnControls 0, 0, 1, 0
    cmdRegister.Enabled = True
    txtCode1.Enabled = False
    txtCode2.Enabled = False
    txtCode3.Enabled = False
    txtCode4.Enabled = False
    txtCode5.Enabled = False
    txtHospitalName.Enabled = False
    StatusBar1.Panels(1).Text = "Cannot find server..."
    If wsCodeServer.State <> 0 Then wsCodeServer.Close
End If
Me.MousePointer = 0
End Sub

Private Sub RetrieveRegistration(mRegCode As String)
Dim Requestheader As String
Dim PayloadData As String
EnControls 0, 0, 0, 0
wsCodeServer.Protocol = sckTCPProtocol
wsCodeServer.RemoteHost = "mdmpi.com.ph"
wsCodeServer.RemotePort = 9998
wsCodeServer.Connect
PayloadData = "INTERFACECODE=" & mRegCode & ""
PayloadData = Blowfish.EncryptString(PayloadData, myKey, True)
Requestheader = "POST /api_interfaceregistration HTTP/1.0" & vbCrLf
Requestheader = Requestheader & "Host: mdmpi.com.ph:9998" & vbCrLf
Requestheader = Requestheader & "Accept-Language: en-US,en;q=0.5" & vbCrLf
'Requestheader = Requestheader & "Accept -Encoding: gzip , deflate" & vbCrLf
Requestheader = Requestheader & "Content-Type: text/plain" & vbCrLf
Requestheader = Requestheader & "Content-Length: " & Len(PayloadData) & vbCrLf
'Requestheader = Requestheader & "Origin: http://mdmpi.com.ph:9998" & vbCrLf
Requestheader = Requestheader & "DNT: 1" & vbCrLf
Requestheader = Requestheader & "Connection: keep -alive" & vbCrLf
'Requestheader = Requestheader & "Referer: http://mdmpi.com.ph:9998/postpage" & vbCrLf
Requestheader = Requestheader & "Upgrade-Insecure-Requests: 1" & vbCrLf
While Not blnConnected And CancelCount < 30
    DoEvents
    Me.MousePointer = 11
    StatusBar1.Panels(1).Text = "Connecting to server(" & CancelCount & ")..."
Wend
If CancelCount >= 30 Then 'And mOpenTemp Then
    lblTempCode.Visible = True
    txtTempCode.Visible = True
End If
CancelCount = 1
If blnConnected And CancelCount < 30 Then
    wsCodeServer.SendData Requestheader & vbCrLf & PayloadData
Else
    EnControls 0, 0, 1, 0
    cmdRegister.Enabled = True
    txtCode1.Enabled = False
    txtCode2.Enabled = False
    txtCode3.Enabled = False
    txtCode4.Enabled = False
    txtCode5.Enabled = False
    txtHospitalName.Enabled = False
    StatusBar1.Panels(1).Text = "Cannot find server..."
    If wsCodeServer.State <> 0 Then wsCodeServer.Close
End If
Me.MousePointer = 0
End Sub

Private Sub wsCodeServer_Close()
    wsCodeServer.Close
    blnConnected = False
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
Dim tempstr As String
Dim arrTemp() As String
Dim StartCtr As Integer
Dim PayloadData As String
Dim tMess As String
tMess = mMess
arrTemp = Split(mMess, " ")
If UBound(arrTemp) > 1 Then
    tempstr = arrTemp(1)
Else
    tempstr = Replace(Mid(tMess, InStr(1, tMess, "HTTP/1.1"), InStr(1, tMess, "OK") - InStr(1, tMess, "HTTP/1.1")), "HTTP/1.1 ", "")
End If
'tMess = Blowfish.DecryptString(mMess, myKey, True)
If Trim(tempstr) = "200" Then
    StatusBar1.Panels(1).Text = "Connected to server..."
    PayloadData = Mid(tMess, InStr(1, tMess, vbCrLf & vbCrLf))
    If Left(PayloadData, 4) = vbCrLf & vbCrLf Then
        PayloadData = Mid(PayloadData, 5)
    End If
    PayloadData = Blowfish.DecryptString(PayloadData, myKey, True)
    If InStr(1, PayloadData, "|") > 0 Then
        arrTemp = Split(PayloadData, "|")
        Select Case arrTemp(0)
            Case "REGERROR"
                StatusBar1.Panels(1).Text = arrTemp(1)
                EnControls 1, 1, 1, 0
            Case "REGOK"
                StatusBar1.Panels(1).Text = arrTemp(1)
                lblRegname.Caption = arrTemp(2)
                lblExpiration.Caption = Format(arrTemp(3), "MMMM dd, yyyy")
                SavePayload PayloadData
                EnControls 0, 0, 0, 1
                SaveSetting "Lab-I", "Registration", "Code1", txtCode1.Text
                SaveSetting "Lab-I", "Registration", "Code2", txtCode2.Text
                SaveSetting "Lab-I", "Registration", "Code3", txtCode3.Text
                SaveSetting "Lab-I", "Registration", "Code4", txtCode4.Text
                SaveSetting "Lab-I", "Registration", "Code5", txtCode5.Text
                SaveSetting "Lab-I", "Registration", "Registered Name", lblRegname.Caption
                SaveSetting "Lab-I", "Registration", "Hospital Name", txtHospitalName.Text
                SaveSetting "Lab-I", "Registration", "Expiration", Format(arrTemp(3), "mm/dd/yyyy")
            Case "REGUPDATE"
                StatusBar1.Panels(1).Text = arrTemp(1)
                EnControls 0, 0, 0, 1
        End Select
    Else
        StatusBar1.Panels(1).Text = "Server Response Error..."
        EnControls 1, 1, 1, 0
    End If
ElseIf Val(Trim(tempstr)) >= "400" Then
    StatusBar1.Panels(1).Text = Trim(tempstr) & "-Client Error..."
    EnControls 1, 1, 1, 0
ElseIf Val(Trim(tempstr)) >= "500" Then
    StatusBar1.Panels(1).Text = Trim(tempstr) & "-Server Error..."
    EnControls 1, 1, 1, 0
Else
    StatusBar1.Panels(1).Text = "Cannot find server..."
End If
End Sub

Private Sub EnControls(enCode As Boolean, enHospName As Boolean, enButton As Boolean, enNewCode As Boolean)
txtCode1.Enabled = enCode
txtCode2.Enabled = enCode
txtCode3.Enabled = enCode
txtCode4.Enabled = enCode
txtCode5.Enabled = enCode
txtHospitalName.Enabled = enHospName
If enButton Then
    If cmdRegister.Caption <> "Retrieve Registration" Then
        cmdRegister.Enabled = ValidateCode
    Else
        cmdRegister.Enabled = enButton
    End If
Else
    cmdRegister.Enabled = enButton
End If
cmdNewCode.Enabled = enNewCode
End Sub

Private Sub SavePayload(mPayloadData As String)
Dim mDirectory As String
Dim mFile As String
Dim iFileNo As Integer
Dim NewPayloadData As String
NewPayloadData = Blowfish.EncryptString(String(50, "-") & mPayloadData & String(50, "-"), myKey, True)
iFileNo = FreeFile
mDirectory = Replace(Format(DBDatabase, ">"), "LABORATORY.MDB", "")
Open mDirectory & "\Laboratory.dat" For Output As #iFileNo
Print #iFileNo, NewPayloadData
Close #iFileNo
End Sub

