VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMLLPClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
' HL7 over TCP (MLLP) sender for middleware mode. Non-blocking: modOutbox.PumpTcp
' calls SendMessage, then polls State every second and calls CheckTimeout.
'   Frame: <VT> message <FS><CR>
'   Accepted when the reply has MSA-1 = AA or CA (and MSA-2 matches, if present).
' It uses an existing Winsock control (frmReceiveData.Winsock1), so no form
' designer change is needed.
'==============================================================================

Private WithEvents ws As MSWinsockLib.Winsock
Attribute ws.VB_VarHelpID = -1

Private mState As Integer
Private mBuffer As String
Private mPayload As String
Private mControlID As String
Private mStarted As Date
Private mTimeout As Long
Private mKeepAlive As Boolean
Private mAccepted As Boolean
Private mLastError As String
Private mHost As String
Private mPort As Long

Public Sub Attach(Sock As MSWinsockLib.Winsock)
    Set ws = Sock
    mState = MLLP_IDLE
End Sub

Public Property Get State() As Integer
    State = mState
End Property

Public Property Get Accepted() As Boolean
    Accepted = mAccepted
End Property

Public Property Get LastError() As String
    LastError = mLastError
End Property

Public Sub Reset()
    mState = MLLP_IDLE
    mBuffer = ""
    mPayload = ""
End Sub

Public Sub Shutdown()
On Error Resume Next
    If Not ws Is Nothing Then
        If ws.State <> sckClosed Then ws.Close
    End If
    Set ws = Nothing
    mState = MLLP_IDLE
End Sub

Public Sub SendMessage(ByVal Payload As String, ByVal ControlID As String, ByVal Host As String, ByVal Port As Long, _
                       ByVal TimeoutSeconds As Long, ByVal KeepAlive As Boolean)
On Error GoTo errFound
    If ws Is Nothing Then
        Finish False, "Winsock not attached"
        Exit Sub
    End If
    mPayload = Payload
    mControlID = ControlID
    mTimeout = TimeoutSeconds
    If mTimeout < 2 Then mTimeout = 2
    mKeepAlive = KeepAlive
    mAccepted = False
    mBuffer = ""
    mStarted = Now
    'reuse an open connection only if it goes to the same LIS
    If ws.State = sckConnected And StrComp(Host, mHost, vbTextCompare) = 0 And Port = mPort Then
        SendFrame
    Else
        If ws.State <> sckClosed Then ws.Close
        mHost = Host
        mPort = Port
        ws.Protocol = sckTCPProtocol
        ws.RemoteHost = Host
        ws.RemotePort = Port
        ws.LocalPort = 0
        mState = MLLP_CONNECTING
        ws.Connect
    End If
    Exit Sub
errFound:
    Finish False, "Connect error: " & Err.Description
End Sub

Private Sub SendFrame()
On Error GoTo errFound
    mBuffer = ""
    mState = MLLP_WAITACK
    mStarted = Now
    ws.SendData Chr(11) & mPayload & Chr(28) & Chr(13)
    Exit Sub
errFound:
    Finish False, "Send error: " & Err.Description
End Sub

'Called every second by the pump while connecting / waiting for the ACK.
Public Sub CheckTimeout()
    If mState = MLLP_CONNECTING Or mState = MLLP_WAITACK Then
        If DateDiff("s", mStarted, Now) > mTimeout Then
            If mState = MLLP_CONNECTING Then
                Finish False, "Timeout connecting to " & mHost & ":" & mPort
            Else
                Finish False, "No ACK from LIS within " & mTimeout & "s"
            End If
        End If
    End If
End Sub

Private Sub Finish(ByVal IsAccepted As Boolean, ByVal ErrorText As String)
On Error Resume Next
    mAccepted = IsAccepted
    mLastError = ErrorText
    If Not IsAccepted Or Not mKeepAlive Then
        If Not ws Is Nothing Then
            If ws.State <> sckClosed Then ws.Close
        End If
    End If
    mState = MLLP_DONE
End Sub

Private Sub ParseAck(ByVal Frame As String)
Dim arrSeg() As String
Dim arrFld() As String
Dim ctr As Long
Dim mCode As String
Dim mText As String
    arrSeg = Split(Replace(Frame, vbLf, ""), vbCr)
    For ctr = 0 To UBound(arrSeg)
        If UCase(Left(Trim(arrSeg(ctr)), 4)) = "MSA|" Then
            arrFld = Split(Trim(arrSeg(ctr)) & "|||", "|")
            mCode = UCase(Trim(arrFld(1)))
            mText = Trim(arrFld(3))
            If mCode = "AA" Or mCode = "CA" Then
                If Len(Trim(arrFld(2))) > 0 And Trim(arrFld(2)) <> mControlID Then
                    Finish False, "ACK for a different message (" & Trim(arrFld(2)) & ")"
                Else
                    Finish True, ""
                End If
            Else
                Finish False, "LIS rejected the message: " & mCode & IIf(Len(mText) > 0, " " & mText, "")
            End If
            Exit Sub
        End If
    Next ctr
    Finish False, "Reply from LIS has no MSA segment"
End Sub

Private Sub ws_Connect()
    If mState = MLLP_CONNECTING Then SendFrame
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
On Error GoTo errFound
Dim sData As String
Dim mStart As Long
Dim mEnd As Long
    ws.GetData sData, vbString
    If mState <> MLLP_WAITACK Then Exit Sub        'unsolicited data: ignore
    mBuffer = mBuffer & sData
    mEnd = InStr(1, mBuffer, Chr(28))
    If mEnd > 0 Then
        mStart = InStr(1, mBuffer, Chr(11))
        If mStart = 0 Or mStart > mEnd Then mStart = 0
        ParseAck Mid(mBuffer, mStart + 1, mEnd - mStart - 1)
        mBuffer = ""
    End If
    Exit Sub
errFound:
    Finish False, "Receive error: " & Err.Description
End Sub

Private Sub ws_Close()
On Error Resume Next
    ws.Close
    If mState = MLLP_CONNECTING Or mState = MLLP_WAITACK Then Finish False, "Connection closed by LIS"
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, _
                     ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    CancelDisplay = True
    If mState = MLLP_CONNECTING Or mState = MLLP_WAITACK Then
        Finish False, "Winsock " & Number & ": " & Description
    Else
        On Error Resume Next
        ws.Close
    End If
End Sub

Private Sub Class_Terminate()
    Shutdown
End Sub
