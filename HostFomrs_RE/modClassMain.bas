Attribute VB_Name = "modClassMain"
Option Explicit

Public Const myKey = "3BlindMice"
Public Blowfish As New clsBlowfish

Public myConnection As New ADODB.Connection
Public extConnection As New ADODB.Connection
Public extRecordset As New ADODB.Recordset
Public rstSettings As New ADODB.Recordset

Public DBDatabase As String
Public DBUserName As String
Public DBPassword As String

Public InstrumentID As String
Public globalSectionID As String
Public strSectionName As String
Public PatientID As String
Public RequestID As String
Public rPatientID As String
Public rRequestID As String
Public strSQLStatement As String
Public strRequestBy As String
Public strRequestByName As String
Public intWorkstation As Integer
Public barCodePrint As Object
Public searchObj As Object
Public isExiting As Boolean

Public Report As New CRAXDRT.Report

Public Sub setFormColors(myForm As Form)
Dim ctr As Integer
On Error Resume Next
    myForm.fraTitle.BackColor = GetSetting("LMS", "FormColors", "TitleBar", &HA66B2B)
    myForm.lblTitle.ForeColor = GetSetting("LMS", "FormColors", "TitleText", &HFFFFFF)
    For ctr = 0 To myForm.fraBody.UBound
        myForm.fraBody(ctr).BackColor = GetSetting("LMS", "FormColors", "BodyFrame", &H6B4B01)
    Next ctr
    myForm.lblStatus.ForeColor = GetSetting("LMS", "FormColors", "BodyText", &HFFFFFF)
    For ctr = 0 To myForm.lblText.UBound
        myForm.lblText(ctr).ForeColor = GetSetting("LMS", "FormColors", "BodyText", &HFFFFFF)
    Next ctr
    myForm.msfgResult.BackColorFixed = GetSetting("LMS", "FormColors", "TitleBar", &HA66B2B)
    myForm.msfgResult.ForeColorFixed = GetSetting("LMS", "FormColors", "TitleText", &HFFFFFF)
    myForm.BackColor = GetSetting("LMS", "FormColors", "BodyFrame", &H6B4B01)
End Sub

Public Sub highLight()
On Error Resume Next
    With Screen.ActiveForm
        If (TypeOf .ActiveControl Is TextBox) Then
            .ActiveControl.SelStart = 0
            .ActiveControl.SelLength = Len(.ActiveControl)
        End If
    End With
End Sub

Public Sub highLight_bg()
On Error Resume Next
    With Screen.ActiveForm
        If (TypeOf .ActiveControl Is TextBox) Then
            If .ActiveControl.Locked Then
                .ActiveControl.BackColor = &H80FF80
            Else
                .ActiveControl.BackColor = &HC0E0FF
            End If
        End If
    End With
End Sub

Public Function showfrmExamSelection(mFilter As String, mUserID As String, Optional mSpecimen As String, Optional mSectionID As String) As String
    
    With frmExamSelection
        .setFilter = mFilter
        .setUserID = mUserID
        .setSpecimen = mSpecimen
        .setSectionID = mSectionID
        .Show 1
        showfrmExamSelection = .setReturn
    End With
    
End Function

Public Function showfrmSelection(mSearchTrans As String, mFilter As String, blnShowFrm As Boolean, _
        Optional mFilter2 As String, Optional mFilter3 As String, Optional mFilter4 As String, _
        Optional mFilter5 As String) As String
    Dim rsSelection As New ADODB.Recordset
    Dim intRSColCtr As Integer
    Dim strReturn As String
    If blnShowFrm Then
        With frmSelection
            .setRS = Nothing
            .setDataTrans = mSearchTrans
            .setFilter = mFilter
            .setFilter2 = mFilter2
            .setFilter3 = mFilter3
            .setFilter4 = mFilter4
            .setFilter5 = mFilter5
            .Show 1
            showfrmSelection = .setReturn
        End With
    Else
        Set rsSelection = myConnection.Execute("EXEC sp_SearchCodes2 " & mSearchTrans & ",'" & mFilter & "','" & _
                mFilter2 & "','" & mFilter3 & "','" & mFilter4 & "','" & mFilter5 & "'")
        With rsSelection
            If .RecordCount = 1 Then
                For intRSColCtr = 0 To .Fields.Count - 1
                    strReturn = strReturn & .Fields(intRSColCtr).Value & "|"
                Next
                strReturn = Mid(strReturn, 1, Len(strReturn) - 1)
                showfrmSelection = strReturn
            ElseIf .RecordCount > 1 Then
                  With frmSelection
                    .setDataTrans = mSearchTrans
                    .setRS = rsSelection
                    .Show 1
                    showfrmSelection = .setReturn
                End With
            End If
        End With
    End If
End Function

Public Function showSummaryRequest(mRID As String, Optional mSectionID As String = "", Optional mSampleID As String = "") As String
        With frmSummaryRequest
        .SetRequestID = mRID
        .setSectionID = mSectionID
        .SetSampleID = mSampleID
        .Show 1
        showSummaryRequest = .toUnload
    End With
End Function

Public Function getGlobalSettings() As String
    Set rstSettings = myConnection.Execute("EXEC sp_LoadCatSetting 'ALL','RE_REQ'")
End Function

Public Sub setGlobalSettings(mSetName As String, mSetValue As String, mSetCategory As String)
    Dim rstSettings As ADODB.Recordset
    Set rstSettings = New ADODB.Recordset
    Set rstSettings = myConnection.Execute("EXEC sp_LoadCatSetting 'ALL','" & mSetCategory & "'")
    With rstSettings
        .Filter = "SetName = '" & mSetName & "'"
        If .RecordCount = 0 Then
            myConnection.Execute "Insert into tblabsettings (InstrumentID,SetName, SetValues, SetCategory) values ('ALL','" & mSetName & "','" & mSetValue & "', '" & mSetCategory & "')"
        Else
            myConnection.Execute "Update tblabsettings SET SetValues = '" & mSetValue & "'" & _
                                    " WHERE InstrumentID = 'ALL' AND SetName = '" & mSetName & "' AND SetCategory = '" & mSetCategory & "'"
        End If
        rstSettings.Requery
    End With
End Sub

Public Sub printBarCode(strBarCode As String, Optional strRequestID As String = "", Optional intNumCopies As String = "1")
    Dim strPatientID As String
    Dim strPHospID As String
    Dim strName As String
    Dim strSection As String
    Dim strDate As String
    Dim strSpecimen As String
    Dim strSampleID As String
    Dim strSID2 As String
    Dim PrintedSampleID As String
    Dim rsBarCode As New ADODB.Recordset
    Dim DataToPrint As String
    Dim ConCatTest As String
    Dim rsBarCodeCtrRow As Integer
    
    
    If IsNumeric(intNumCopies) Then
        
        If barCodePrint Is Nothing Then
            Set barCodePrint = CreateObject("BCSET.Barcode")
        End If
        
        If Len(Trim(strRequestID)) = 0 Then
            If InStr(1, strBarCode, "-") > 0 Then
                Set rsBarCode = myConnection.Execute("EXEC sp_getRVSampleIDToPrint_Select '" & Mid(strBarCode, 1, InStr(1, strBarCode, "-") - 1) & "',NULL")
            Else
                Set rsBarCode = myConnection.Execute("EXEC sp_getRVSampleIDToPrint_Select '" & strBarCode & "',NULL")
            End If
            If rsBarCode.RecordCount > 0 Then
                With rsBarCode
                    For rsBarCodeCtrRow = 0 To .RecordCount - 1
                        If rsBarCodeCtrRow = 0 Then
                            
                            DataToPrint = .Fields("PID").Value & "|" & .Fields("PHOSTPID").Value & "|" & .Fields("PLNAME").Value & _
                            "|" & .Fields("PFNAME").Value & "|" & .Fields("PMNAME").Value & "|" & .Fields("PDISPLAYNAME").Value & "|" & .Fields("PGENDER").Value & _
                            "|" & .Fields("PDOB").Value & "|" & .Fields("RequestID").Value & "|" & Format(.Fields("RequestDate").Value, "mm/dd/yyyy") & _
                            "|" & Format(.Fields("RequestDate").Value, "HH:mm") & "|" & .Fields("RequestLocation").Value & _
                            "|" & .Fields("RequestHospReqID").Value & "|" & .Fields("RequestWard").Value & "|" & .Fields("RequestRoom").Value & _
                            "|" & .Fields("RequestAge").Value & "|" & .Fields("RequestPriority").Value & "|" & .Fields("RequestPType").Value & _
                            "|" & .Fields("Reserve1").Value & "|" & .Fields("Reserve2").Value & "|" & .Fields("Reserve3").Value & _
                            "|" & .Fields("Reserve4").Value & "|" & .Fields("Reserve5").Value & "|" & .Fields("RequestPhysicianID").Value & _
                            "|" & .Fields("RequestExamRateClassID").Value & "|" & .Fields("RequestBed").Value & "|" & .Fields("RequestComment").Value & _
                            "|" & strBarCode & "|" & Right(.Fields("RVSampleID").Value, 3) & "|" & "<CONCATTESTSREQ>" & "|" & .Fields("Specimen").Value & _
                            "|" & .Fields("SECTIONCODE") & "|" & .Fields("SECTIONNAME").Value & "|" & .Fields("SID2").Value
                            ConCatTest = .Fields("ExamRequestName").Value
                            strSection = .Fields("SECTIONNAME").Value
                        Else
                            ConCatTest = ConCatTest & "," & .Fields("ExamRequestName").Value
                        End If
                        .MoveNext
                    Next
                End With
                DataToPrint = Replace(DataToPrint, "<CONCATTESTSREQ>", ConCatTest)
                barCodePrint.printBarCode DataToPrint, strSection, intNumCopies
                
'                With rsBarCode
'                    strPatientID = .Fields("PID")
'                    strName = .Fields("PLNAME").Value & ", " & .Fields("PFNAME").Value & " " & IIf(IsNull(.Fields("PMNAME").Value), "", .Fields("PMNAME").Value)
'                    strSection = .Fields("SECTIONNAME").Value
'                    strDate = Format(.Fields("RequestDate").Value, "MM/dd/yyyy HH:mm")
'                    strSpecimen = .Fields("Specimen").Value 'IIf(IsNull(.Fields("SID2")), .Fields("Specimen").Value, .Fields("SID2"))
'                    strSampleID = IIf(InStr(1, strBarCode, "-") > 0, strBarCode, .Fields("RVSampleID").Value)
'                    'strSID2 = IIf(IsNull(.Fields("SID2")), "", .Fields("SID2"))
'                    strPHospID = IIf(IsNull(.Fields("PHOSTPID")), "", .Fields("PHOSTPID"))
'                End With
'
'                If strBarCode = strSID2 Then
'                    barCodePrint.printBarCode strPatientID, strName, strSampleID, _
'                                            strSection, Right(Trim(strSampleID), 4), _
'                                            strDate, intNumCopies, strSpecimen, strBarCode, 2, strPHospID
'                Else
'                    barCodePrint.printBarCode strPatientID, strName, strSampleID, _
'                                            strSection, Right(Trim(strSampleID), 4), _
'                                            strDate, intNumCopies, strSpecimen, , , strPHospID
'                End If
            End If
            rsBarCode.Close
        Else
            If Len(Trim(strBarCode)) > 0 And strBarCode <> "NEWPATIENT" Then
                Set rsBarCode = myConnection.Execute("EXEC sp_getRVSampleIDToPrint_Select '" & strBarCode & "','" & strRequestID & "'")
            Else
                Set rsBarCode = myConnection.Execute("EXEC sp_getRVSampleIDToPrint_Select NULL,'" & strRequestID & "'")
            End If
            If rsBarCode.RecordCount > 0 Then
                With rsBarCode
                    For rsBarCodeCtrRow = 0 To .RecordCount - 1
                        If Len(strSampleID) = 0 Then
                            strSampleID = .Fields("RVSampleID")
                            DataToPrint = .Fields("PID").Value & "|" & .Fields("PHOSTPID").Value & "|" & .Fields("PLNAME").Value & _
                            "|" & .Fields("PFNAME").Value & "|" & .Fields("PMNAME").Value & "|" & .Fields("PDISPLAYNAME").Value & "|" & .Fields("PGENDER").Value & _
                            "|" & .Fields("PDOB").Value & "|" & .Fields("RequestID").Value & "|" & Format(.Fields("RequestDate").Value, "mm/dd/yyyy") & _
                            "|" & Format(.Fields("RequestDate").Value, "HH:mm") & "|" & .Fields("RequestLocation").Value & _
                            "|" & .Fields("RequestHospReqID").Value & "|" & .Fields("RequestWard").Value & "|" & .Fields("RequestRoom").Value & _
                            "|" & .Fields("RequestAge").Value & "|" & .Fields("RequestPriority").Value & "|" & .Fields("RequestPType").Value & _
                            "|" & .Fields("Reserve1").Value & "|" & .Fields("Reserve2").Value & "|" & .Fields("Reserve3").Value & _
                            "|" & .Fields("Reserve4").Value & "|" & .Fields("Reserve5").Value & "|" & .Fields("RequestPhysicianID").Value & _
                            "|" & .Fields("RequestExamRateClassID").Value & "|" & .Fields("RequestBed").Value & "|" & .Fields("RequestComment").Value & _
                            "|" & .Fields("RVSampleID").Value & "|" & Right(.Fields("RVSampleID").Value, 3) & "|" & "<CONCATTESTSREQ>" & "|" & .Fields("Specimen").Value & _
                            "|" & .Fields("SECTIONCODE") & "|" & .Fields("SECTIONNAME").Value & "|" & .Fields("SID2").Value
                            ConCatTest = .Fields("ExamRequestName").Value
                            strSection = .Fields("SECTIONNAME").Value
                        Else
                            ConCatTest = ConCatTest & "," & .Fields("ExamRequestName").Value
                        End If
                        .MoveNext
                        If .EOF = True Then
                            DataToPrint = Replace(DataToPrint, "<CONCATTESTSREQ>", ConCatTest)
                            barCodePrint.printBarCode DataToPrint, strSection, CStr(intNumCopies)
                            'pID As String, pName As String, pSID As String, pSection As String, pSeq As String, pDate As String, pCopies As String, Optional pSpecimen As String = "", Optional pSID2 As String = "", Optional pSIDTrans As Integer = 1, Optional pHospID As String = ""
                            
                        Else
                            If strSampleID <> .Fields("RVSampleID") Then
                                DataToPrint = Replace(DataToPrint, "<CONCATTESTSREQ>", ConCatTest)
                                barCodePrint.printBarCode DataToPrint, strSection, intNumCopies
                                DataToPrint = ""
                                strSampleID = ""
                            End If
                        End If
                    Next
                End With
            End If
            rsBarCode.Close
            
            
'            If rsBarCode.RecordCount > 0 Then
'                 With rsBarCode
'                    Do
'                        strPatientID = .Fields("PID")
'                        strName = .Fields("PLNAME").Value & ", " & .Fields("PFNAME").Value & " " & IIf(IsNull(.Fields("PMNAME").Value), "", .Fields("PMNAME").Value)
'                        strSection = .Fields("SECTIONNAME").Value
'                        strDate = Format(.Fields("RequestDate").Value, "MM/dd/yyyy HH:mm")
'                        strSpecimen = .Fields("Specimen").Value
'                        strSampleID = .Fields("RVSampleID").Value
'                        strSID2 = IIf(IsNull(.Fields("SID2")), "", .Fields("SID2"))
'                        strPHospID = IIf(IsNull(.Fields("PHOSTPID")), "", .Fields("PHOSTPID"))
'
'                        If Len(Trim(strSID2)) > 0 Then
'                            If Len(PrintedSampleID) = 0 Or InStr(1, PrintedSampleID, strSID2) = 0 Then
'                                barCodePrint.printBarCode strPatientID, strName, strSampleID, _
'                                                    strSection, Right(Trim(strSampleID), 4), _
'                                                    strDate, intNumCopies, strSpecimen, strSID2, 2, strPHospID
'                                PrintedSampleID = PrintedSampleID & "|" & strSID2
'                            End If
'                        Else
'                            barCodePrint.printBarCode strPatientID, strName, strSampleID, _
'                                                    strSection, Right(Trim(strSampleID), 4), _
'                                                    strDate, intNumCopies, strSpecimen, , , strPHospID
''                            If Len(PrintedSampleID) = 0 Or InStr(1, PrintedSampleID, strSampleID) = 0 Then
''                                barCodePrint.printBarCode strPatientID, strName, strSampleID, _
''                                                    strSection, Right(Trim(strSampleID), 4), _
''                                                    strDate, intNumCopies, strSpecimen, , , strPHospID
''                                PrintedSampleID = PrintedSampleID & "|" & strSampleID
''                            End If
'                        End If
'                        .MoveNext
'                    Loop Until .EOF
'                 End With
'            End If
'
        End If
'        Debug.Print strPatientID & "|" & strName & "|" & strSection & "|" & strDate & "|" & strSpecimen & "|" & strSampleID & "|" & strSID2
    End If
End Sub

Public Function getValue(TextToSearch As String, ColNumber As Integer) As String
    Dim strTemp() As String

    strTemp = Split(TextToSearch, "|")

    getValue = strTemp(ColNumber)
End Function


'Public Sub AdditionalWorks(strWorkID As String, ParamArray strParameters())
'
'    Select Case strWorkID
'        Case "1"
'            Dim rsHosts As New ADODB.Recordset
'            Dim strHosts As String
'            Dim RowNum As Integer
'            Dim TempCtr As Integer
'
'            Set rsHosts = myConnection.Execute("EXEC sp_GetReqExamHosts_Select " & "'" & strParameters(1) & "','" & _
'                                                        strParameters(0) & "'")
'
'            strHosts = rsHosts.Fields(0)
'            For TempCtr = 1 To Len(strHosts)
'                If Mid(strHosts, TempCtr, 1) = "|" Then
'                    RowNum = RowNum + 1
'                End If
'            Next
'
'            If Len(Trim(strHosts)) > 0 Then
'                If strParameters(4) = "<PROFILE>" Then
'                For TempCtr = 1 To RowNum
'                    Set rsHosts = myConnection.Execute("EXEC sp_tbLabInstrumentRequest2_Insert '" & strParameters(0) & "','" & _
'                                                            getValue(TempCtr, strHosts) & "','" & strParameters(2) & "','" & _
'                                                            strParameters(3) & "','0','" & strParameters(4) & "'")
'                Next
'            End If
'        Case "2"
'            Dim rsDelInsReqTwo As New ADODB.Recordset
'
'            Set rsDelInsReqTwo = myConnection.Execute("sp_tbLabInstrumentRequest2_delete " & "'" & strParameters(0) & "','" & _
'                                                        strParameters(1) & "','" & strParameters(2) & "','" & strParameters(3) & "'")
'    End Select
'End Sub

Public Function getUserDet(strUserID As String) As String
    Dim rsUserDet As New ADODB.Recordset
    Dim strToReturn As String
    Dim LabUserColCtr As Integer
    
    strSQLStatement = "SELECT UserRights,UserDisplay,UserLName,UserFName,UserMName FROM tbLabUsers WHERE UserID =  '" & strUserID & "'"
    Set rsUserDet = myConnection.Execute(strSQLStatement)
    
    If rsUserDet.RecordCount > 0 Then
        For LabUserColCtr = 0 To rsUserDet.Fields.Count - 1
            If Len(strToReturn) = 0 Then
                strToReturn = Trim(rsUserDet.Fields(LabUserColCtr))
            Else
                strToReturn = strToReturn & "|" & Trim(rsUserDet.Fields(LabUserColCtr))
            End If
        Next
    End If
    
    getUserDet = strToReturn
    
End Function

Public Function CompAge(date1 As Date, date2 As Date) As String
    
    Dim tempComputedAge As String
    Dim diffYear As String
    Dim diffMonth As String
    Dim diffDay As String
    
    diffYear = DateDiff("yyyy", date2, date1)
    diffMonth = DateDiff("m", date2, date1)
    diffDay = DateDiff("d", date2, date1)

    CompAge = diffYear & " " & diffMonth & " " & diffDay
        
    If Val(diffYear) = 1 Then
        If Val(Month(date2)) = 12 Then
            If Val(Abs(diffDay)) > 0 Then
                If Abs(Val(diffDay)) > 1 Then
                    tempComputedAge = Val(Abs(diffDay)) & " DAYS"
                Else
                    tempComputedAge = Val(Abs(diffDay)) & " DAY"
                End If
            Else
                tempComputedAge = "NEWBORN"
            End If
        Else
            If Val(Month(date1)) = Val(Month(date2)) Then
                If Val(Day(date1)) >= Val(Day(date2)) Then
                    tempComputedAge = diffYear
                Else
                    If Val(diffMonth) > 1 Then
                        tempComputedAge = diffMonth & " MONTHS"
                    Else
                        tempComputedAge = diffMonth & " MONTH"
                    End If
                End If
            ElseIf Val(Month(date1)) < Val(Month(date2)) Then
                If Val(diffMonth) > 1 Then
                    tempComputedAge = diffMonth & " MONTHS"
                Else
                    tempComputedAge = diffMonth & " MONTH"
                End If
            Else
                tempComputedAge = diffYear
            End If
        End If
    ElseIf Val(diffYear) > 1 Then
        If Val(Month(date1)) = Val(Month(date2)) Then
            If Val(Day(date1)) >= Val(Day(date2)) Then
                tempComputedAge = diffYear
            Else
                tempComputedAge = diffYear - 1
            End If
        ElseIf Val(Month(date1)) < Val(Month(date2)) Then
            tempComputedAge = diffYear - 1
        Else
            tempComputedAge = diffYear
        End If
    Else
        If Val(diffYear) = 0 Then
            If Val(diffMonth) > 0 Then
                If Val(diffMonth) > 1 Then
                    tempComputedAge = diffMonth & " MONTHS"
                Else
                    If Val(diffDay) > 0 And Val(diffDay) < 31 Then
                        If Abs(Val(diffDay)) > 1 Then
                            tempComputedAge = Val(diffDay) & " DAYS"
                        Else
                            tempComputedAge = Val(diffDay) & " DAY"
                        End If
                    Else
                        tempComputedAge = diffMonth & " MONTH"
                    End If
                End If
            Else
                If Val(Abs(diffDay)) > 0 Then
                    If Abs(Val(diffDay)) > 1 Then
                        tempComputedAge = Val(Abs(diffDay)) & " DAYS"
                    Else
                        tempComputedAge = Val(Abs(diffDay)) & " DAY"
                    End If
                Else
                    tempComputedAge = "NEWBORN"
                End If
            End If
        Else
            If Val(diffMonth) = 12 And (Day(date1) >= Day(date2)) Then
                tempComputedAge = "1"
            Else
                diffYear = Val(diffYear) - 1
            
                If Val(diffYear) = 0 Then
                    tempComputedAge = diffMonth & " MONTHS"
                End If
            End If
        End If
    End If

    CompAge = tempComputedAge

End Function

Public Function FindSetting(mSetName As String, mDefault As String) As String

    On Error GoTo errfound
        If rstSettings.RecordCount > 0 Then
            rstSettings.Filter = "SetName = '" & mSetName & "'"
            If rstSettings.RecordCount > 0 Then
                FindSetting = rstSettings.Fields("SetValues").Value
            Else
                FindSetting = mDefault
            End If
            rstSettings.Filter = ""
        Else
            FindSetting = mDefault
        End If

errfound:
        If Err.Number = 3021 Then
            FindSetting = mDefault
        ElseIf Err.Number <> 0 Then
            FindSetting = mDefault
        End If

End Function

Public Function FindSetting2(mSetName As String, msfgSettings As MSHFlexGrid, mDefault As String) As String
    
    Dim ctrMsfgSettings As Integer
    Dim returnSetting As String
    With msfgSettings
        For ctrMsfgSettings = 0 To .Rows - 1
            If mSetName = UCase(.TextMatrix(ctrMsfgSettings, 0)) Then
                returnSetting = .TextMatrix(ctrMsfgSettings, 1)
                msfgSettings.RemoveItem (ctrMsfgSettings)
                Exit For
            End If
        Next ctrMsfgSettings
    End With
          
    If Len(returnSetting) = 0 Then
        returnSetting = mDefault
    End If
    
    FindSetting2 = returnSetting

End Function

Public Function OpenMainReport(ByVal strReportFileName As String, ParamArray strParameters()) As String

    Dim crxParameterField As CRAXDRT.ParameterFieldDefinition
    
    Dim intCtr As Integer
    Dim intTotalParam As Integer
    Dim crxtable As CRAXDRT.DatabaseTable
    Dim crxApplication As New CRAXDRT.Application
    On Error GoTo errfound
    
    intTotalParam = UBound(strParameters)
    
    Set Report = crxApplication.OpenReport(strReportFileName, 1)
    For Each crxtable In Report.Database.Tables
       crxtable.SetLogOnInfo "Laboratory", DBDatabase, DBUserName, DBPassword
    Next
    
    If intTotalParam >= 0 Then
       For Each crxParameterField In Report.ParameterFields
          If intTotalParam >= intCtr Then
             crxParameterField.AddCurrentValue strParameters(intCtr)
          Else
             Exit For
          End If
          intCtr = intCtr + 1
       Next
    End If

    Report.DisplayProgressDialog = True
    Report.PrintOut False
        
errfound:
    If Err.Number = -2147206460 Then
        'MsgBox "Result format not found. Please setup a Request format first.", vbOKOnly + vbExclamation, "Request Form"
        OpenMainReport = "Error"
    ElseIf Err.Number = -2147206461 Then
        'MsgBox "Report File Not Found." & vbCrLf & vbCrLf & strReportFileName, vbExclamation + vbOKOnly, "Report"
        OpenMainReport = "Error"
    ElseIf Err.Number <> 0 Then
        'MsgBox Err.Number & ": " & Err.Description, vbOKOnly + vbExclamation, "Result Form"
        OpenMainReport = "Error"
    Else
        OpenMainReport = ""
    End If
    Set Report = Nothing
    Set crxParameterField = Nothing
    Set crxtable = Nothing
    Set crxApplication = Nothing
End Function
