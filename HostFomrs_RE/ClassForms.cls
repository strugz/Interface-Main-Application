VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Let wConnection(ByVal vNewValue As ADODB.Connection)
    Set myConnection = vNewValue
    DBDatabase = myConnection.Properties(69)
    DBUserName = myConnection.Properties(68)
    DBPassword = myConnection.Properties(66)
    getGlobalSettings
End Property

Public Property Let wInstrumentID(ByVal vNewValue As String)
    InstrumentID = vNewValue
End Property

Public Property Let wSectionName(ByVal vNewValue As String)
    Dim rsSection As New ADODB.Recordset
    Set rsSection = myConnection.Execute("EXEC sp_SearchCodes 2, NULL ,'" & vNewValue & "'")
    
    If rsSection.RecordCount > 0 Then
        strSectionName = rsSection.Fields(1)
    End If
End Property

Public Property Let SearchRequestFormLeft(ByVal vNewValue As Long)
    frmSearchRequest.Left = vNewValue
End Property

Public Property Let SearchRequestFormTop(ByVal vNewValue As Long)
    frmSearchRequest.Top = vNewValue
End Property

Public Property Let setWorkStationID(ByVal vNewValue As Integer)
    intWorkstation = vNewValue
End Property

Public Function showSearchForm(strUserID As String, strUserName As String, strSectionID As String, Optional strSearchstring As String = "") As String
    Dim errload
    strRequestBy = strUserID
    strRequestByName = strUserName
    globalSectionID = strSectionID
    rPatientID = ""
    rRequestID = ""
    'frmSearchRequest.wSearchString = strSearchstring
    On Error GoTo errload
    frmSearchRequest.Show 1
    isExiting = False
    showSearchForm = rPatientID & "|" & rRequestID
errload:
    Resume Next
End Function

Public Sub showRequestSettings()
    frmRequestSetting.Show 1
    isExiting = False
End Sub

Public Sub showWardRoomsSettings()
    frmWardsRooms.Show 1
    isExiting = False
End Sub

Public Function EditRequest(strPatientID As String, strRequestID As String, strSectionID As String, _
                strSpecimen As String, strSampleNumber As String, strUserID As String, strUserName As String) As String
    With frmEditRequest
        isExiting = False
        rPatientID = ""
        rRequestID = ""
        .setPatientID = strPatientID
        .SetRequestID = strRequestID
        .setSectionID = strSectionID
        .setSpecimen = strSpecimen
        .setSampleNumber = strSampleNumber
        .setUserID = strUserID
        .Show 1
        EditRequest = rPatientID & "|" & rRequestID
    End With
End Function

Public Function CheckInSample(strRID As String, strSectionID As String, strSpecimen As String, strUserID As String) As String
    Dim strCheckIn As String
    Dim strDate As String
    Dim strSampleID As String
    Dim rsCheckIn As New ADODB.Recordset
    Dim intCtr As Integer
    
    strCheckIn = "EXEC sp_CheckRequestID_select '" & strRID & "'"
    Set rsCheckIn = myConnection.Execute(strCheckIn)

    If rsCheckIn.RecordCount > 0 Then
        strDate = Format(rsCheckIn.Fields("RequestDate").Value, "MM/dd/yyyy")
        
        strSQLStatement = "EXEC sp_CheckInBySmpl_update '" & strRID & "','" & strSectionID & "'," & _
                            "'" & strSpecimen & "','" & strDate & "'," & strUserID
                            
        Set rsCheckIn = myConnection.Execute(strSQLStatement)
        If rsCheckIn.RecordCount > 0 Then
            strSampleID = rsCheckIn.Fields("SampleID").Value
            modClassMain.printBarCode strSampleID, , 1
        Else
            strSampleID = ""
        End If
    End If
    
    CheckInSample = strSampleID
    
End Function

Public Function newTest(strPatientID As String, strRequestID As String, strUserID As String, strUserName As String) As String
'    With frmReqExam
'        PatientID = strPatientID
'        RequestID = strRequestID
'        strRequestBy = strUserID
'        strRequestByName = strUserName
'        .Show 1
'    End With
End Function

Public Sub printBarCode(strBarCodeNum As String, Optional NumCopy As String = "1")
    If IsNumeric(Val(NumCopy)) Then
        modClassMain.printBarCode strBarCodeNum, , Int(NumCopy)
    End If
End Sub

Public Sub printRequestForm(strRequestID As String, strUserID As String)
    If Len(Trim(strRequestID)) > 0 And Len(Trim(strUserID)) > 0 Then
        OpenMainReport App.Path & "\Reports\MSSQL SERVER\Request.rpt", strRequestID, strUserID, "0"
    End If
End Sub

Public Sub showReqSummary(strRID As String, Optional strSampleID As String = "")
    showSummaryRequest strRID, , strSampleID
End Sub

Public Sub showReqRates()
    With frmReqRates
        .Show 1
    End With
End Sub

Public Sub DelReq(strRID As String)
    With frmDelReq
        .Show 1
    End With
End Sub

Public Sub CloseForms()
    Dim ctr As Integer
    Dim tempStr As String
    Dim mhWnd() As Form
    
    If Not searchObj Is Nothing Then
        searchObj.CloseForms
    End If
    
    ctr = 0
    tempStr = ""
    isExiting = True
    Dim frm As Form
    For Each frm In Forms
        ctr = ctr + 1
    Next frm
    ReDim mhWnd(ctr) As Form
    ctr = 0
    For Each frm In Forms
        ctr = ctr + 1
        Set mhWnd(ctr) = frm
    Next frm
    Do While ctr <> 0
         Unload mhWnd(ctr)
         Set mhWnd(ctr) = Nothing
         ctr = ctr - 1
    Loop

End Sub

Public Property Get Version() As String
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Public Function getDispTestReqSum() As String
    getDispTestReqSum = FindSetting("DISPTESTREQSUM", "0")
End Function

Private Sub Class_Initialize()
    Set searchObj = CreateObject("SearchCodes.ClassSearch")
End Sub
