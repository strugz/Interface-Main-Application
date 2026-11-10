Attribute VB_Name = "modOutbox"
'==============================================================================
' Middleware mode: store-and-forward outbox (see docs\MIDDLEWARE_MODE.md)
'
' Results are turned into HL7 ORU^R01 messages and written to a file spool:
'   <OutboxRoot>\<Destination>\pending\<ControlID>.msg
'   <OutboxRoot>\<Destination>\sent\     (delivered and acknowledged)
'   <OutboxRoot>\<Destination>\failed\   (gave up after MaxAttempts; RESEND to retry)
' Destinations: HL7TCP (MLLP to the LIS/HIS) and HL7FILE (file drop folder).
'
' Spool file layout: one header line, CRLF, then the HL7 payload (CR separated).
'   #LABI-OUTBOX|1|InstrumentID|SampleNo|ControlID|Attempts|NextTry|LastError
'
' Registry: Lab-I\Middleware (Enable, HL7Tcp, HL7TcpHost, HL7TcpPort, AckTimeout,
' KeepAlive, HL7File, HL7FileExt, HL7FileCRLF, RetrySeconds, MaxAttempts,
' KeepSentDays, OutboxDir, SendUnmapped, SendingFacility, ReceivingFacility).
'==============================================================================
Option Explicit

Public Const MW_DEST_TCP As String = "HL7TCP"
Public Const MW_DEST_FILE As String = "HL7FILE"

'clsMLLPClient states
Public Const MLLP_IDLE As Integer = 0
Public Const MLLP_CONNECTING As Integer = 1
Public Const MLLP_WAITACK As Integer = 2
Public Const MLLP_DONE As Integer = 3

Private Const MW_HEADER As String = "#LABI-OUTBOX"

'Set by frmMiddleware when the user asks to switch mode; the caller then runs
'mdiMain.SwitchModeWithRestart after the modal screen has closed.
Public MW_SwitchRequested As Boolean

Private mMLLP As clsMLLPClient
Private mCurrentTcpFile As String
Private mPumpBusy As Boolean
Private mSeq As Long
Private mLastPrune As String
Private mLastStatus As Single
Private mExamMap As Object          'Scripting.Dictionary: InstExamName -> Array(HisExamCode, DisplayName, ExamNotes)
Private mFso As Object

'------------------------------------------------------------------------------
' Settings
'------------------------------------------------------------------------------
Public Function MW_Setting(KeyName As String, Optional DefaultValue As String = "") As String
    MW_Setting = GetSetting("Lab-I", "Middleware", KeyName, DefaultValue)
End Function

Public Function MW_TcpEnabled() As Boolean
    MW_TcpEnabled = (Val(MW_Setting("HL7Tcp", "0")) = 1)
End Function

Public Function MW_FileEnabled() As Boolean
    MW_FileEnabled = (Val(MW_Setting("HL7File", "0")) = 1)
End Function

Private Function Fso() As Object
    If mFso Is Nothing Then Set mFso = CreateObject("Scripting.FileSystemObject")
    Set Fso = mFso
End Function

Public Function OutboxRoot() As String
Dim mRoot As String
    mRoot = Trim(MW_Setting("OutboxDir", ""))
    If Len(mRoot) = 0 Then
        If DBServerType = "LOCAL DATABASE" Then
            mRoot = Replace(DBDatabase, "Laboratory.mdb", "", , , vbTextCompare) & "Outbox"
        Else
            mRoot = App.Path & "\Outbox"
        End If
    End If
    If Right(mRoot, 1) = "\" Then mRoot = Left(mRoot, Len(mRoot) - 1)
    OutboxRoot = mRoot
End Function

Private Function DestFolder(Dest As String, SubFolder As String) As String
    DestFolder = OutboxRoot & "\" & Dest & "\" & SubFolder
End Function

Private Sub EnsureFolder(ByVal mPath As String)
Dim mParent As String
    If Fso.FolderExists(mPath) Then Exit Sub
    mParent = Fso.GetParentFolderName(mPath)
    If Len(mParent) > 0 Then EnsureFolder mParent
    Fso.CreateFolder mPath
End Sub

Private Sub EnsureDestFolders(Dest As String)
    EnsureFolder DestFolder(Dest, "pending")
    EnsureFolder DestFolder(Dest, "sent")
    EnsureFolder DestFolder(Dest, "failed")
End Sub

'------------------------------------------------------------------------------
' Start / stop (called from frmReceiveData Form_Load / Form_Unload)
'------------------------------------------------------------------------------
Public Sub OutboxStart(Sock As MSWinsockLib.Winsock)
On Error GoTo errFound
    Set mExamMap = Nothing          'reload the test code map on every connect
    mCurrentTcpFile = ""
    If Not MiddlewareMode Then Exit Sub
    Set mMLLP = New clsMLLPClient
    mMLLP.Attach Sock
    If MW_TcpEnabled Then EnsureDestFolders MW_DEST_TCP
    If MW_FileEnabled Then EnsureDestFolders MW_DEST_FILE
    Exit Sub
errFound:
    Debug.Print "OutboxStart: " & Err.Description
End Sub

Public Sub OutboxStop()
On Error Resume Next
    If Not mMLLP Is Nothing Then mMLLP.Shutdown
    Set mMLLP = Nothing
    mCurrentTcpFile = ""
End Sub

'------------------------------------------------------------------------------
' Enqueue: called from frmReceiveData.processData with the plugin contract fields.
' obxAll: OBX lines (vbCr separated), one block per patient joined by "<SP>".
' oriAll: ORI lines, one per patient, joined by vbCrLf.
' Returns a line for the receive window ("" when there is nothing to report).
'------------------------------------------------------------------------------
Public Function OutboxEnqueueResult(ByVal obxAll As String, ByVal oriAll As String) As String
Dim arrBlocks() As String
Dim arrOri() As String
Dim oriLines() As String
Dim nOri As Long
Dim ctr As Long
Dim mORI As String
Dim mPayload As String
Dim mControlID As String
Dim arrFields() As String
Dim mSample As String
Dim mQueued As Long
Dim mSkipped As Long
Dim mResult As String
    If Not MiddlewareMode Then Exit Function
    If Not (MW_TcpEnabled Or MW_FileEnabled) Then
        OutboxEnqueueResult = "Outbox: no destination enabled (Lab-I\Middleware\HL7Tcp / HL7File)."
        Exit Function
    End If
    If Len(Trim(obxAll)) = 0 Or UCase(Trim(obxAll)) = "ERRCONV" Or UCase(Trim(oriAll)) = "ERRCONV" Then Exit Function

    'Collect the non-empty ORI lines
    arrOri = Split(Replace(oriAll, vbLf, ""), vbCr)
    ReDim oriLines(0)
    nOri = 0
    For ctr = 0 To UBound(arrOri)
        If Len(Trim(arrOri(ctr))) > 0 Then
            ReDim Preserve oriLines(nOri)
            oriLines(nOri) = arrOri(ctr)
            nOri = nOri + 1
        End If
    Next ctr

    arrBlocks = Split(obxAll, "<SP>")
    For ctr = 0 To UBound(arrBlocks)
        If Len(Trim(Replace(Replace(arrBlocks(ctr), vbCr, ""), vbLf, ""))) > 0 Then
            If nOri = 0 Then
                mORI = ""
            ElseIf ctr < nOri Then
                mORI = oriLines(ctr)
            Else
                mORI = oriLines(nOri - 1)
            End If
            mControlID = NextControlID
            mPayload = MW_BuildORU(arrBlocks(ctr), mORI, mControlID)
            arrFields = Split(mORI & "||||", "|")
            mSample = arrFields(3)
            If Len(mPayload) = 0 Then
                mSkipped = mSkipped + 1
            Else
                If MW_TcpEnabled Then WriteSpool MW_DEST_TCP, mControlID, mSample, mPayload
                If MW_FileEnabled Then WriteSpool MW_DEST_FILE, mControlID, mSample, mPayload
                mQueued = mQueued + 1
                If Len(mResult) > 0 Then mResult = mResult & ", "
                mResult = mResult & IIf(Len(Trim(mSample)) > 0, Trim(mSample), "(no sample no.)")
            End If
        End If
    Next ctr

    If mQueued > 0 Then OutboxEnqueueResult = "Outbox: queued " & mQueued & " result message(s) for LIS/HIS: " & mResult
    If mSkipped > 0 Then
        If Len(OutboxEnqueueResult) > 0 Then OutboxEnqueueResult = OutboxEnqueueResult & vbCrLf
        OutboxEnqueueResult = OutboxEnqueueResult & "Outbox: " & mSkipped & " result block(s) had no mapped tests and were not queued."
    End If
End Function

'Keeps a raw message that could not be queued (plugin or build error), so it can be handled by hand.
Public Sub OutboxSaveUnqueued(ByVal RawMessage As String, ByVal Reason As String)
On Error Resume Next
Dim mFolder As String
Dim ts As Object
    mFolder = OutboxRoot & "\unqueued"
    EnsureFolder mFolder
    Set ts = Fso.OpenTextFile(mFolder & "\" & InstrumentID & "_" & Format(Now, "yyyymmdd_hhnnss") & ".txt", 8, True)
    ts.WriteLine "===== " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & Reason & " ====="
    ts.WriteLine RawMessage
    ts.Close
End Sub

Private Function NextControlID() As String
    mSeq = (mSeq + 1) Mod 10000
    NextControlID = Format(Now, "yymmddhhnnss") & Format(mSeq, "0000")
End Function

Private Sub WriteSpool(Dest As String, ControlID As String, SampleNo As String, Payload As String)
Dim mFile As String
Dim ts As Object
    EnsureDestFolders Dest
    mFile = DestFolder(Dest, "pending") & "\" & ControlID & ".msg"
    Set ts = Fso.CreateTextFile(mFile & ".tmp", True)
    ts.Write BuildHeader(InstrumentID, SampleNo, ControlID, 0, "", "") & vbCrLf & Payload
    ts.Close
    Fso.MoveFile mFile & ".tmp", mFile
End Sub

Private Function CleanField(ByVal s As String) As String
    s = Replace(s, "|", "/")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    CleanField = s
End Function

Private Function BuildHeader(InstID As String, SampleNo As String, ControlID As String, Attempts As Long, NextTry As String, LastError As String) As String
    BuildHeader = MW_HEADER & "|1|" & CleanField(InstID) & "|" & CleanField(SampleNo) & "|" & CleanField(ControlID) & "|" & Attempts & "|" & CleanField(NextTry) & "|" & CleanField(Left(LastError, 200))
End Function

'Reads a spool file. Returns False when it cannot be read.
Private Function ReadSpool(mFile As String, Header() As String, Payload As String) As Boolean
On Error GoTo errFound
Dim ts As Object
Dim mText As String
Dim mPos As Long
    Set ts = Fso.OpenTextFile(mFile, 1, False)
    If ts.AtEndOfStream Then mText = "" Else mText = ts.ReadAll
    ts.Close
    mPos = InStr(1, mText, vbCrLf)
    If Left(mText, Len(MW_HEADER)) <> MW_HEADER Or mPos = 0 Then
        Header = Split(MW_HEADER & "|1||||0||", "|")
        Payload = mText
    Else
        Header = Split(Left(mText, mPos - 1) & "||||||||", "|")
        Payload = Mid(mText, mPos + 2)
    End If
    ReadSpool = True
    Exit Function
errFound:
    ReadSpool = False
End Function

Private Sub RewriteSpool(mFile As String, Header() As String, Payload As String)
Dim ts As Object
    Set ts = Fso.CreateTextFile(mFile & ".tmp", True)
    ts.Write BuildHeader(Header(2), Header(3), Header(4), CLng(Val(Header(5))), Header(6), Header(7)) & vbCrLf & Payload
    ts.Close
    If Fso.FileExists(mFile) Then Fso.DeleteFile mFile, True
    Fso.MoveFile mFile & ".tmp", mFile
End Sub

'Oldest pending spool file (names start with the timestamp), or "".
Private Function OldestPending(Dest As String) As String
Dim mName As String
Dim mBest As String
    mName = Dir(DestFolder(Dest, "pending") & "\*.msg")
    Do While Len(mName) > 0
        If Len(mBest) = 0 Or StrComp(mName, mBest, vbTextCompare) < 0 Then mBest = mName
        mName = Dir
    Loop
    If Len(mBest) > 0 Then OldestPending = DestFolder(Dest, "pending") & "\" & mBest
End Function

Private Function IsDue(Header() As String) As Boolean
    If Len(Trim(Header(6))) <> 14 Then
        IsDue = True
    Else
        IsDue = (Format(Now, "yyyymmddhhnnss") >= Header(6))
    End If
End Function

Private Sub MoveSpool(mFile As String, Dest As String, SubFolder As String)
Dim mTarget As String
    EnsureFolder DestFolder(Dest, SubFolder)
    mTarget = DestFolder(Dest, SubFolder) & "\" & Fso.GetFileName(mFile)
    If Fso.FileExists(mTarget) Then Fso.DeleteFile mTarget, True
    Fso.MoveFile mFile, mTarget
End Sub

'Delivered: move to sent\. Failed: count the attempt, back off, or move to failed\.
Private Sub CompleteSpool(mFile As String, Dest As String, Accepted As Boolean, ErrorText As String)
Dim Header() As String
Dim Payload As String
Dim mAttempts As Long
Dim mWait As Long
Dim mMax As Long
    If Not ReadSpool(mFile, Header, Payload) Then Exit Sub
    mAttempts = CLng(Val(Header(5))) + 1
    Header(5) = CStr(mAttempts)
    If Accepted Then
        Header(6) = ""
        Header(7) = "OK " & Format(Now, "yyyy-mm-dd hh:nn:ss")
        RewriteSpool mFile, Header, Payload
        MoveSpool mFile, Dest, "sent"
    Else
        Header(7) = Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & ErrorText
        mMax = CLng(Val(MW_Setting("MaxAttempts", "10")))
        If mMax < 1 Then mMax = 1
        If mAttempts >= mMax Then
            Header(6) = ""
            RewriteSpool mFile, Header, Payload
            MoveSpool mFile, Dest, "failed"
            frmReceiveData.OutboxNotify "Outbox " & Dest & ": message " & Header(4) & " (sample " & Header(3) & ") FAILED after " & mAttempts & " attempts: " & ErrorText
        Else
            mWait = CLng(Val(MW_Setting("RetrySeconds", "30")))
            If mWait < 5 Then mWait = 5
            If mAttempts > 1 Then mWait = mWait * (2 ^ IIf(mAttempts - 1 > 5, 5, mAttempts - 1))
            If mWait > 600 Then mWait = 600
            Header(6) = Format(DateAdd("s", mWait, Now), "yyyymmddhhnnss")
            RewriteSpool mFile, Header, Payload
            frmReceiveData.OutboxNotify "Outbox " & Dest & ": send failed (" & ErrorText & "), retry " & mAttempts & "/" & mMax & " in " & mWait & "s"
        End If
    End If
End Sub

'------------------------------------------------------------------------------
' Pump: called every second from frmReceiveData.tmrStatus_Timer
'------------------------------------------------------------------------------
Public Sub OutboxPump()
    If Not MiddlewareMode Then Exit Sub
    If mPumpBusy Then Exit Sub
    mPumpBusy = True
    On Error GoTo errFound
    If MW_TcpEnabled Then PumpTcp
    If MW_FileEnabled Then PumpFile
    If mLastPrune <> Format(Now, "yyyymmdd") Then
        mLastPrune = Format(Now, "yyyymmdd")
        PruneSent
    End If
    If Abs(Timer - mLastStatus) >= 5 Then
        mLastStatus = Timer
        ShowStatus
    End If
    mPumpBusy = False
    Exit Sub
errFound:
    mPumpBusy = False
    Debug.Print "OutboxPump: " & Err.Description
End Sub

Private Sub PumpTcp()
Dim mFile As String
Dim Header() As String
Dim Payload As String
    If mMLLP Is Nothing Then Exit Sub
    Select Case mMLLP.State
        Case MLLP_IDLE
            mFile = OldestPending(MW_DEST_TCP)
            If Len(mFile) = 0 Then Exit Sub
            If Not ReadSpool(mFile, Header, Payload) Then Exit Sub
            If Not IsDue(Header) Then Exit Sub        'keep HL7 order: wait for the oldest one
            If Len(Trim(MW_Setting("HL7TcpHost", ""))) = 0 Or Val(MW_Setting("HL7TcpPort", "0")) <= 0 Then
                CompleteSpool mFile, MW_DEST_TCP, False, "HL7TcpHost/HL7TcpPort not set"
                Exit Sub
            End If
            mCurrentTcpFile = mFile
            mMLLP.SendMessage Payload, Header(4), Trim(MW_Setting("HL7TcpHost", "")), CLng(Val(MW_Setting("HL7TcpPort", "0"))), _
                CLng(Val(MW_Setting("AckTimeout", "10"))), (Val(MW_Setting("KeepAlive", "1")) = 1)
        Case MLLP_DONE
            If Len(mCurrentTcpFile) > 0 Then
                If Fso.FileExists(mCurrentTcpFile) Then CompleteSpool mCurrentTcpFile, MW_DEST_TCP, mMLLP.Accepted, mMLLP.LastError
            End If
            mCurrentTcpFile = ""
            mMLLP.Reset
        Case Else
            mMLLP.CheckTimeout
    End Select
End Sub

Private Sub PumpFile()
Dim mFile As String
Dim Header() As String
Dim Payload As String
Dim mDir As String
Dim mExt As String
Dim mTarget As String
Dim ts As Object
Dim ctr As Integer
    mDir = GetSetting("LAB-i", "HL7 Settings", "FWDIR", App.Path & "\HostSpooler")
    mExt = Trim(MW_Setting("HL7FileExt", "hl7"))
    For ctr = 1 To 20                       'at most 20 files per tick
        mFile = OldestPending(MW_DEST_FILE)
        If Len(mFile) = 0 Then Exit Sub
        If Not ReadSpool(mFile, Header, Payload) Then Exit Sub
        If Not IsDue(Header) Then Exit Sub
        On Error GoTo errWrite
        If Not Fso.FolderExists(mDir) Then Err.Raise 76, , "Folder not found: " & mDir
        mTarget = mDir & "\" & Header(4)
        If Len(mExt) > 0 Then mTarget = mTarget & "." & mExt
        If Val(MW_Setting("HL7FileCRLF", "1")) = 1 Then Payload = Replace(Replace(Payload, vbCrLf, vbCr), vbCr, vbCrLf)
        'write under a temp name, then rename, so the LIS never reads a half-written file
        Set ts = Fso.CreateTextFile(mTarget & ".part", True)
        ts.Write Payload
        ts.Close
        If Fso.FileExists(mTarget) Then Fso.DeleteFile mTarget, True
        Fso.MoveFile mTarget & ".part", mTarget
        On Error GoTo 0
        CompleteSpool mFile, MW_DEST_FILE, True, ""
    Next ctr
    Exit Sub
errWrite:
    Dim mErr As String
    mErr = Err.Description
    On Error GoTo 0
    CompleteSpool mFile, MW_DEST_FILE, False, mErr
End Sub

'Deletes sent\ files older than KeepSentDays (failed\ is never pruned).
Private Sub PruneSent()
On Error Resume Next
Dim mDays As Long
Dim mDest As Variant
Dim fil As Object
    mDays = CLng(Val(MW_Setting("KeepSentDays", "90")))
    If mDays <= 0 Then Exit Sub
    For Each mDest In Array(MW_DEST_TCP, MW_DEST_FILE)
        If Fso.FolderExists(DestFolder(CStr(mDest), "sent")) Then
            For Each fil In Fso.GetFolder(DestFolder(CStr(mDest), "sent")).Files
                If DateDiff("d", fil.DateLastModified, Now) > mDays Then fil.Delete True
            Next fil
        End If
    Next mDest
End Sub

'------------------------------------------------------------------------------
' Status and console helpers
'------------------------------------------------------------------------------
Private Function CountFiles(Dest As String, SubFolder As String) As Long
On Error Resume Next
    CountFiles = 0
    If Fso.FolderExists(DestFolder(Dest, SubFolder)) Then CountFiles = Fso.GetFolder(DestFolder(Dest, SubFolder)).Files.Count
End Function

Public Function OutboxStatusShort() As String
Dim s As String
    If MW_TcpEnabled Then s = "TCP " & CountFiles(MW_DEST_TCP, "pending") & " pend/" & CountFiles(MW_DEST_TCP, "failed") & " fail"
    If MW_FileEnabled Then
        If Len(s) > 0 Then s = s & "  "
        s = s & "File " & CountFiles(MW_DEST_FILE, "pending") & " pend/" & CountFiles(MW_DEST_FILE, "failed") & " fail"
    End If
    If Len(s) = 0 Then s = "no destination"
    OutboxStatusShort = "Outbox: " & s
End Function

Private Sub ShowStatus()
On Error Resume Next
    With mdiMain.sbrStatus.Panels(7)
        .Visible = True
        .AutoSize = sbrContents
        .Text = OutboxStatusShort
    End With
End Sub

Public Function OutboxStatusText() As String
Dim s As String
Dim mDest As Variant
    s = "MIDDLEWARE MODE: " & IIf(MiddlewareMode, "ON", "OFF") & vbCrLf & "OUTBOX FOLDER: " & OutboxRoot
    For Each mDest In Array(MW_DEST_TCP, MW_DEST_FILE)
        s = s & vbCrLf & mDest & ": " & IIf(IIf(mDest = MW_DEST_TCP, MW_TcpEnabled, MW_FileEnabled), "enabled", "disabled") & _
            "  pending " & CountFiles(CStr(mDest), "pending") & ", sent " & CountFiles(CStr(mDest), "sent") & ", failed " & CountFiles(CStr(mDest), "failed")
    Next mDest
    If MW_TcpEnabled Then
        s = s & vbCrLf & "LIS: " & MW_Setting("HL7TcpHost", "(not set)") & ":" & MW_Setting("HL7TcpPort", "(not set)")
        If Not mMLLP Is Nothing Then
            If Len(mMLLP.LastError) > 0 Then s = s & "  last error: " & mMLLP.LastError
        End If
    End If
    If MW_FileEnabled Then s = s & vbCrLf & "FILE DROP: " & GetSetting("LAB-i", "HL7 Settings", "FWDIR", App.Path & "\HostSpooler")
    OutboxStatusText = s
End Function

'Moves failed\ messages back to pending\ with the attempt count reset. Dest "" = all.
Public Function OutboxResend(Optional ByVal Dest As String = "") As Long
On Error GoTo errFound
Dim mDest As Variant
Dim fil As Object
Dim colFiles As Collection
Dim mFile As Variant
Dim Header() As String
Dim Payload As String
Dim mCount As Long
    For Each mDest In Array(MW_DEST_TCP, MW_DEST_FILE)
        If Len(Dest) = 0 Or UCase(Dest) = mDest Then
            If Fso.FolderExists(DestFolder(CStr(mDest), "failed")) Then
                Set colFiles = New Collection
                For Each fil In Fso.GetFolder(DestFolder(CStr(mDest), "failed")).Files
                    colFiles.Add fil.Path
                Next fil
                For Each mFile In colFiles
                    If OutboxResendFile(CStr(mFile), CStr(mDest)) Then mCount = mCount + 1
                Next mFile
            End If
        End If
    Next mDest
    OutboxResend = mCount
    Exit Function
errFound:
    OutboxResend = mCount
End Function

'--- helpers for the Middleware screen (frmMiddleware) -----------------------

Public Function OutboxFolder(ByVal Dest As String, ByVal SubFolder As String) As String
    OutboxFolder = DestFolder(Dest, SubFolder)
End Function

Public Function OutboxCount(ByVal Dest As String, ByVal SubFolder As String) As Long
    OutboxCount = CountFiles(Dest, SubFolder)
End Function

Public Function OutboxReadSpool(ByVal mFile As String, Header() As String, Payload As String) As Boolean
    OutboxReadSpool = ReadSpool(mFile, Header, Payload)
End Function

'Puts one message back in the queue with its attempt count reset. A message that is
'already pending is only made due now (it is not moved onto itself).
Public Function OutboxResendFile(ByVal mFile As String, ByVal Dest As String) As Boolean
On Error GoTo errFound
Dim Header() As String
Dim Payload As String
    If Not Fso.FileExists(mFile) Then Exit Function
    If Not ReadSpool(mFile, Header, Payload) Then Exit Function
    Header(5) = "0"
    Header(6) = ""
    Header(7) = "Resent " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    RewriteSpool mFile, Header, Payload
    If StrComp(Fso.GetParentFolderName(mFile), DestFolder(Dest, "pending"), vbTextCompare) <> 0 Then
        MoveSpool mFile, Dest, "pending"
    End If
    OutboxResendFile = True
    Exit Function
errFound:
    OutboxResendFile = False
End Function

'Spool files of one folder, newest first (names start with the yymmddhhnnss control ID).
'TotalCount returns how many files there are in all.
Public Function OutboxListFiles(ByVal Dest As String, ByVal SubFolder As String, ByVal MaxItems As Long, TotalCount As Long) As Collection
Dim col As New Collection
Dim arrNames() As String
Dim mName As String
Dim n As Long
Dim i As Long
Dim j As Long
Dim mGap As Long
Dim mTemp As String
    TotalCount = 0
    Set OutboxListFiles = col
    If Not Fso.FolderExists(DestFolder(Dest, SubFolder)) Then Exit Function
    ReDim arrNames(0 To 63)
    mName = Dir(DestFolder(Dest, SubFolder) & "\*.msg")
    Do While Len(mName) > 0
        If n > UBound(arrNames) Then ReDim Preserve arrNames(0 To n * 2)
        arrNames(n) = mName
        n = n + 1
        mName = Dir
    Loop
    TotalCount = n
    If n = 0 Then Exit Function
    'shell sort, descending
    mGap = n \ 2
    Do While mGap > 0
        For i = mGap To n - 1
            mTemp = arrNames(i)
            j = i
            Do While j >= mGap
                If StrComp(arrNames(j - mGap), mTemp, vbTextCompare) >= 0 Then Exit Do
                arrNames(j) = arrNames(j - mGap)
                j = j - mGap
            Loop
            arrNames(j) = mTemp
        Next i
        mGap = mGap \ 2
    Loop
    For i = 0 To n - 1
        If i >= MaxItems Then Exit For
        col.Add DestFolder(Dest, SubFolder) & "\" & arrNames(i)
    Next i
End Function

'------------------------------------------------------------------------------
' HL7 ORU^R01 builder from the plugin contract fields (works for every plugin)
'------------------------------------------------------------------------------
Public Function HL7Escape(ByVal s As String) As String
    s = Replace(s, "\", "\E\")
    s = Replace(s, "|", "\F\")
    s = Replace(s, "^", "\S\")
    s = Replace(s, "~", "\R\")
    s = Replace(s, "&", "\T\")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    HL7Escape = Trim(s)
End Function

Private Function Fld(arr() As String, Index As Long) As String
    If Index <= UBound(arr) Then Fld = Trim(arr(Index)) Else Fld = ""
End Function

Private Sub LoadExamMap()
On Error GoTo errFound
Dim rst As ADODB.Recordset
Dim mKey As String
    Set mExamMap = CreateObject("Scripting.Dictionary")
    mExamMap.CompareMode = 1        'text compare
    If Not isDBconnected Then Exit Sub
    Set rst = myConnection.Execute("select InstExamName, HisExamCode, DisplayName, ExamNotes from tbcolabexam where InstrumentID = '" & Replace(InstrumentID, "'", "''") & "'")
    Do While Not rst.EOF
        mKey = Trim(rst.Fields("InstExamName").Value & "")
        If Len(mKey) > 0 And Not mExamMap.Exists(mKey) Then
            mExamMap.Add mKey, Array(Trim(rst.Fields("HisExamCode").Value & ""), Trim(rst.Fields("DisplayName").Value & ""), Trim(rst.Fields("ExamNotes").Value & ""))
        End If
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    Exit Sub
errFound:
    Debug.Print "LoadExamMap: " & Err.Description
End Sub

Private Function HL7Date(ByVal s As String) As String
    s = Trim(s)
    If Len(s) >= 8 And IsNumeric(Left(s, 8)) Then
        HL7Date = Left(s, 8)
    ElseIf IsDate(s) Then
        HL7Date = Format(CDate(s), "yyyymmdd")
    Else
        HL7Date = ""
    End If
End Function

Public Function MW_BuildORU(ByVal obxBlock As String, ByVal oriLine As String, ByVal ControlID As String) As String
Dim ori() As String
Dim arrLines() As String
Dim arrOBX() As String
Dim seg(25) As String
Dim ctr As Long
Dim n As Long
Dim mLine As String
Dim mCode As String
Dim mCodeText As String
Dim mValue As String
Dim mRange As String
Dim mNotes As String
Dim arrNotes() As String
Dim mMapped As Variant
Dim mObsDate As String
Dim mSex As String
Dim mName As String
Dim mFamily As String
Dim mGiven As String
Dim mPID As String
Dim strOBX As String
Dim strMsg As String
Dim i As Integer
    If mExamMap Is Nothing Then LoadExamMap
    ori = Split(oriLine & "|||||||||||||||||||", "|")

    'OBX segments
    arrLines = Split(Replace(obxBlock, vbLf, ""), vbCr)
    n = 0
    For ctr = 0 To UBound(arrLines)
        mLine = Trim(arrLines(ctr))
        If UCase(Left(mLine, 4)) = "OBX|" Then
            arrOBX = Split(mLine, "|")
            mCode = Fld(arrOBX, 3)
            mValue = Fld(arrOBX, 5)
            mMapped = Empty
            If Len(mCode) > 0 And Not mExamMap Is Nothing Then
                If mExamMap.Exists(mCode) Then
                    mMapped = mExamMap(mCode)
                    If mMapped(0) = "" Or mMapped(0) = "0" Then mMapped = Empty
                End If
            End If
            If IsEmpty(mMapped) Then
                If Val(MW_Setting("SendUnmapped", "1")) = 1 Then
                    mCodeText = HL7Escape(mCode) & "^" & HL7Escape(mCode) & "^L"
                    mNotes = ""
                Else
                    mCodeText = ""
                End If
            Else
                mCodeText = HL7Escape(CStr(mMapped(0))) & "^" & HL7Escape(CStr(mMapped(1))) & "^L"
                mNotes = CStr(mMapped(2))
            End If
            If Len(mCodeText) > 0 And Len(mCode) > 0 Then
                'reference range from ExamNotes: "male|female" or one range
                mRange = ""
                If Len(mNotes) > 0 Then
                    If InStr(1, mNotes, "|") > 0 Then
                        arrNotes = Split(mNotes, "|")
                        If UCase(Left(Fld(ori, 11), 1)) = "M" Then
                            mRange = arrNotes(0)
                        ElseIf UCase(Left(Fld(ori, 11), 1)) = "F" Then
                            mRange = arrNotes(1)
                        End If
                    Else
                        mRange = mNotes
                    End If
                End If
                n = n + 1
                For i = 0 To 25: seg(i) = "": Next i
                seg(0) = "OBX"
                seg(1) = CStr(n)
                seg(2) = IIf(IsNumeric(mValue), "NM", "ST")
                seg(3) = mCodeText
                seg(5) = HL7Escape(mValue)
                seg(6) = HL7Escape(Fld(arrOBX, 6))
                seg(7) = HL7Escape(mRange)
                seg(8) = HL7Escape(Replace(Fld(arrOBX, 8), Chr(160), ""))
                seg(11) = "F"
                seg(14) = Fld(arrOBX, 14)
                seg(18) = HL7Escape(Fld(ori, 1))
                If Len(mObsDate) = 0 Then mObsDate = Fld(arrOBX, 14)
                strOBX = strOBX & JoinSeg(seg, 18) & vbCr
            End If
        End If
    Next ctr
    If n = 0 Then Exit Function

    If Len(mObsDate) = 0 Then mObsDate = Format(Now, "yyyymmddhhnnss")

    'MSH
    For i = 0 To 25: seg(i) = "": Next i
    seg(0) = "MSH"
    seg(1) = "^~\&"
    seg(2) = HL7Escape(GetSetting("LAB-i", "HL7 Settings", "Sending Application", "LAB-Iv" & App.Major))
    seg(3) = HL7Escape(MW_Setting("SendingFacility", ""))
    seg(4) = HL7Escape(GetSetting("LAB-i", "HL7 Settings", "Receiving Application", "MDMPI"))
    seg(5) = HL7Escape(MW_Setting("ReceivingFacility", ""))
    seg(6) = Format(Now, "yyyymmddhhnnss")
    seg(8) = "ORU^R01"
    seg(9) = ControlID
    seg(10) = "P"
    seg(11) = HL7Escape(GetSetting("LAB-i", "HL7 Settings", "HL7 Version", "2.5.1"))
    strMsg = JoinSeg(seg, 11) & vbCr        'seg(n) = MSH-(n+1): MSH-1 is the "|" itself

    'PID: patient ID from ORI(7), else the sample number
    mPID = Fld(ori, 7)
    If Len(mPID) = 0 Then mPID = Fld(ori, 3)
    mName = Fld(ori, 2)
    If InStr(1, mName, ",") > 0 Then
        mFamily = Trim(Left(mName, InStr(1, mName, ",") - 1))
        mGiven = Trim(Mid(mName, InStr(1, mName, ",") + 1))
    Else
        mFamily = mName
        mGiven = ""
    End If
    mSex = UCase(Left(Fld(ori, 11), 1))
    If mSex <> "M" And mSex <> "F" And mSex <> "" Then mSex = "U"
    For i = 0 To 25: seg(i) = "": Next i
    seg(0) = "PID"
    seg(1) = "1"
    seg(3) = HL7Escape(mPID)
    seg(5) = HL7Escape(mFamily) & IIf(Len(mGiven) > 0, "^" & HL7Escape(mGiven), "")
    seg(7) = HL7Date(Fld(ori, 8))
    seg(8) = mSex
    strMsg = strMsg & JoinSeg(seg, 8) & vbCr

    'OBR: placer = ORI(6) sequence/accession, filler = ORI(3) sample number
    For i = 0 To 25: seg(i) = "": Next i
    seg(0) = "OBR"
    seg(1) = "1"
    seg(2) = HL7Escape(Fld(ori, 6))
    seg(3) = HL7Escape(Fld(ori, 3))
    seg(4) = HL7Escape(SectionCode) & "^" & HL7Escape(InstrumentID) & "^L"
    seg(7) = mObsDate
    seg(15) = HL7Escape(Fld(ori, 12))
    seg(16) = HL7Escape(Fld(ori, 14))
    seg(22) = Format(Now, "yyyymmddhhnnss")
    seg(24) = HL7Escape(SectionCode)
    seg(25) = "F"
    strMsg = strMsg & JoinSeg(seg, 25) & vbCr

    MW_BuildORU = strMsg & strOBX
End Function

Private Function JoinSeg(seg() As String, LastField As Long) As String
Dim i As Long
Dim s As String
    s = seg(0)
    For i = 1 To LastField
        s = s & "|" & seg(i)
    Next i
    'drop empty trailing fields
    Do While Right(s, 1) = "|"
        s = Left(s, Len(s) - 1)
    Loop
    JoinSeg = s
End Function
