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
Private rstData As ADODB.Recordset

Public LastsearchID As String
Public LastSearchName As String

Public Function showResultEntry(mInstrumentID As String, mPatientName As String, mResultDate As String, mSequenceNumber As String, mMultiStart As Long, mMultiTotal As Long, mMultiCurrent As Long, Optional mPrifileID As String) As String
With frmResultEntry
    .wInstrumentID = mInstrumentID
    .wPatientName = mPatientName
    .wResultdate = mResultDate
    .wSequenceNumber = mSequenceNumber
    .wMultiStart = mMultiStart
    .wMultiTotal = mMultiTotal
    .wMultiCurrent = mMultiCurrent
    .Show 1
    showResultEntry = ReturnValueInt
End With
End Function

Public Property Let wConnection(ByVal vNewValue As ADODB.Connection)
Dim rstTemp As ADODB.Recordset
Dim conTemp As New ADODB.Connection
Dim ctr As Integer
Dim hasLicense As Boolean
Dim hasLicExpire As Boolean
Dim hasPatLicense As Boolean
Dim hasPatLicExpire As Boolean
Dim hasResultFlag As Boolean
Dim hasRemarks2 As Boolean
Dim hasExamType As Boolean
Dim ModifyIRName As Boolean
Dim ModifyReportsTable As Boolean
Dim ModifyReportName As Boolean
Dim ModifyReportSize As Boolean
Dim tempDB As String
Dim tempDBIntrumentResult As String
Dim qryInsert As String
Dim qryUpdate As String
Dim qryReportInsert As String
Dim qryReportUpdate As String
Dim qryUpdateExam As String
Dim qryInsertExam As String
Dim qryDisplayProfileExams As String
Dim qryVSValSelect As String
Dim qryPreviousResult As String
Dim ErrLog As String
Dim ModifyWardSize As Boolean
Dim ModifyPhysicianSize As Boolean
Dim hasExamTypeValues As Boolean
Dim ModifyMTSize As Boolean
On Error GoTo errFound
Set myConnection = vNewValue
Set rstPhysician = myConnection.Execute("select * from tbPhysician")
Set rstMedtech = myConnection.Execute("select * from tbmedtech")
Set rstHospital = myConnection.Execute("select * from tbHospitalInfo where hospitalid = '1'")
ErrLog = ""
hasLicense = False
hasLicExpire = False
hasPatLicense = False
hasPatLicExpire = False
hasResultFlag = False
hasRemarks2 = False
hasExamType = False
hasExamTypeValues = False
ModifyIRName = False
ModifyReportsTable = False
ModifyReportName = False
ModifyReportSize = False
ModifyWardSize = False
ModifyMTSize = False
If GetSetting("Lab-I", "Settings", "UpdatedTable", "0") = "0" Then
    
    qryInsert = "PARAMETERS mRefNum Text ( 255 ), mLastName Text ( 255 ), mFirstName Text ( 255 ), mMiddleName Text ( 255 ), mAge Text ( 255 ), " & _
    "mSex Text ( 255 ), mExamID Text ( 255 ), mFormType Text ( 255 ), mCassNum Text ( 255 ), mRackNum Text ( 255 ), " & _
    "mBirthDate Text ( 255 ), mWard Text ( 255 ), mRoom Text ( 255 ), mResultDate DateTime, mRequestDate DateTime, " & _
    "mPrintDate DateTime, mPhysician Text ( 255 ), mMedTech Text ( 255 ), mProfileID Text ( 255 ), mClassID Text ( 255 ), " & _
    "mPrinted Bit, mInstrumentID Text ( 255 ), mInstrumentResultName Text ( 255 ), mRemarks Text ( 255 ), mProfileName Text ( 255 ), " & _
    "mClassName Text ( 255 ), mPathologist Text ( 255 ), mPatientID Text ( 255 ), mSpecimenID Text ( 255 ), " & _
    "mAdditionalProfiles Text ( 255 ), mCivilStatus Text ( 255 ), mSpecimen Text ( 255 ), mResultFlag Text ( 255 )," & _
    "mMTLic Text ( 255 )='', mMTLicExp Text ( 255 )='', mPTLic Text ( 255 )='', mPTLicExp Text ( 255 )='', mRemarks2 Text ( 255 ); " & _
    "INSERT INTO tbPatientMaster ( RefNum, LastName, FirstName, MiddleName, Age, Sex, ExamID, FormType, CassNum, " & _
    "RackNum, BirthDate, Ward, Room, ResultDate, RequestDate, PrintDate, Physician, MedTech, ProfileID, ClassID, " & _
    "Printed, InstrumentID, InstrumentResultName, Remarks, ProfileName, ClassName, Pathologist, PatientID, SpecimenID, " & _
    "AdditionalProfiles, CivilStatus, Specimen, ResultFlag, MTLicense, MTLicExpire, PatLicense, PatLicExpire, Remarks2) " & _
    "SELECT mRefNum, mLastName, mFirstName, mMiddleName, mAge, mSex,mExamID, mFormType, mCassNum, mRackNum, mBirthDate, mWard,  " & _
    "mRoom, mResultDate, mRequestDate, mPrintDate, mPhysician, mMedTech, mProfileID, mClassID, mPrinted, mInstrumentID, mInstrumentResultName,  " & _
    "mRemarks, mProfileName, mClassName, mPathologist, mPatientID, mSpecimenID, mAdditionalProfiles, mCivilStatus, mSpecimen,mResultFlag, mMTLic, mMTLicExp, mPTLic, mPTLicExp, mRemarks2; "
    
    qryUpdate = "PARAMETERS mRefNum Text ( 255 ), mLastName Text ( 255 ), mFirstName Text ( 255 ), mMiddleName Text ( 255 ), mAge Text ( 255 ),  " & _
    "mSex Text ( 255 ), mExamID Text ( 255 ), mFormType Text ( 255 ), mCassNum Text ( 255 ), mRackNum Text ( 255 ),  " & _
    "mBirthDate Text ( 255 ), mWard Text ( 255 ), mRoom Text ( 255 ), mResultDate DateTime, mRequestDate DateTime,  " & _
    "mPhysician Text ( 255 ), mMedTech Text ( 255 ), mProfileID Text ( 255 ), mClassID Text ( 255 ),  " & _
    "mInstrumentID Text ( 255 ), mInstrumentResultName Text ( 255 ), mRemarks Text ( 255 ), mProfileName Text ( 255 ),  " & _
    "mClassName Text ( 255 ), mPathologist Text ( 255 ), mPatientID Text ( 255 ), mSpecimenID Text ( 255 ),  " & _
    "mAdditionalProfiles Text ( 255 ), mCivilStatus Text ( 255 ), mSpecimen Text ( 255 ), mResultFlag Text ( 255 ), " & _
    "mMTLic Text ( 255 )='', mMTLicExp Text ( 255 )='', mPTLic Text ( 255 )='', mPTLicExp Text ( 255 )='', mRemarks2 Text ( 255 ); " & _
    "UPDATE tbPatientMaster SET tbPatientMaster.LastName = mLastName, tbPatientMaster.FirstName = mFirstName,  " & _
    "tbPatientMaster.MiddleName = mMiddleName, tbPatientMaster.Age = mAge, tbPatientMaster.Sex = mSex, tbPatientMaster.ExamID = mExamID,  " & _
    "tbPatientMaster.FormType = mFormType, tbPatientMaster.CassNum = mCassNum, tbPatientMaster.RackNum = mRackNum,  " & _
    "tbPatientMaster.BirthDate = mBirthDate, tbPatientMaster.Ward = mWard, tbPatientMaster.Room = mRoom,  " & _
    "tbPatientMaster.ResultDate = mResultDate, tbPatientMaster.RequestDate = mRequestDate, tbPatientMaster.Physician = mPhysician,  " & _
    "tbPatientMaster.MedTech = mMedTech, tbPatientMaster.ProfileID = mProfileID, tbPatientMaster.ClassID = mClassID,  " & _
    "tbPatientMaster.InstrumentID = mInstrumentID, tbPatientMaster.InstrumentResultName = mInstrumentResultName,  " & _
    "tbPatientMaster.Remarks = mRemarks, tbPatientMaster.ProfileName = mProfileName, tbPatientMaster.ClassName = mClassName,  " & _
    "tbPatientMaster.Pathologist = mPathologist, tbPatientMaster.PatientID = mPatientID, tbPatientMaster.SpecimenID = mSpecimenID,  " & _
    "tbPatientMaster.AdditionalProfiles = mAdditionalProfiles, tbPatientMaster.CivilStatus = mCivilStatus, tbPatientMaster.Specimen = mSpecimen, tbPatientMaster.ResultFlag = mResultFlag, " & _
    "tbPatientMaster.MTLicense=mMTLic, tbPatientMaster.MTLicExpire=mMTLicExp, tbPatientMaster.PatLicense=mPTLic, tbPatientMaster.PatLicExpire=mPTLicExp, tbPatientMaster.Remarks2 = mRemarks2 " & _
    "WHERE (((tbPatientMaster.[RefNum])=[mRefNum])); "

    qryReportInsert = "PARAMETERS mRptName Text (50), mRptDesc Text (100), mRptSectionCode Text (10), mRptFile OLEOBJECT, mRptUser Text (25), mRptFileSize Long;" & _
    "INSERT INTO tbLabReportNames(rptName, rptDescription, rptDefault, rptSectionCode, rptFile, rptModifyDate, rptModifyUser, rptFileSize)" & _
    "SELECT mRptName, mRptDesc, 0, mRptSectionCode, mRptFile, Date() & ' ' & Time(), mRptUser, mRptFileSize;"
    
    qryReportUpdate = "PARAMETERS mRptID Long, mRptName Text (50), mRptDesc Text (100), mRptFile OLEOBJECT, mRptUser Text (25), mRptFileSize Long; " & _
    "UPDATE tbLabReportNames set rptName = mRptName, rptDescription = mRptDesc, " & _
    "rptFile = mRptFile, rptModifyDate = Date() & ' ' & Time(), rptModifyUser = mRptUser, rptFileSize = mRptFileSize " & _
    "WHERE rptId = mRptID;"
    
    qryUpdateExam = "PARAMETERS mInstrumentID Text ( 255 ), mExamRefNum Text ( 255 ), mInstExamName Text ( 255 ), mHISExamCode Text ( 255 ), " & _
    "mDisplayName Text ( 255 ), mActive Bit, mHISResultCode Text ( 255 ), mUnit Text ( 255 ), mDiff Bit, mDilution Text ( 255 ), " & _
    "mFactorOP Text ( 255 ), mFormula Text ( 255 ), mHiloFlags Text ( 255 ), mExamNotes Text ( 255 ), mExamType Text ( 255 ), mExamTypeValues Text ( 255 );" & _
    "UPDATE tbCoLabExam SET tbCoLabExam.InstExamName = mInstExamName, tbCoLabExam.HISExamCode = mHISExamCode, " & _
    "tbCoLabExam.DisplayName = mDisplayName, tbCoLabExam.Active = mActive, tbCoLabExam.HISResultCode = mHISResultCode, " & _
    "tbCoLabExam.Unit = mUnit, tbCoLabExam.DiffCount = mDiff, tbCoLabExam.Dilution = mDilution, tbCoLabExam.FactorOP = mFactorOP, " & _
    "tbCoLabExam.Formula = mFormula, tbCoLabExam.HiloFlags = mHiloFlags, tbCoLabExam.ExamNotes = mExamNotes, " & _
    "tbCoLabExam.ExamType = mExamType, tbCoLabExam.ExamTypeValues = mExamTypeValues " & _
    "WHERE (((tbCoLabExam.[InstrumentID])=[mInstrumentID]) AND ((tbCoLabExam.[ExamID])=[mExamRefNum]));"

    qryInsertExam = "PARAMETERS mInstrumentID Text ( 255 ), mExamRefNum Text ( 255 ), mInstExamName Text ( 255 ), mHISExamCode Text ( 255 ), " & _
    "mDisplayName Text ( 255 ), mActive Bit, mHISResultCode Text ( 255 ), mUnit Text ( 255 ), mDiff Bit, " & _
    "mDilution Text ( 255 ), mFactorOP Text ( 255 ), mFormula Text ( 255 ), mHiloFlags Text ( 255 ), mExamNotes Text ( 255 ), " & _
    "mExamType Text ( 255 ), mExamTypeValues Text ( 255 );" & _
    "INSERT INTO tbCoLabExam ( InstrumentID, ExamID, InstExamName, HISExamCode, DisplayName, Active, HISResultCode, Unit, " & _
    "DIFFCOUNT, Dilution, FactorOP, Formula, HiloFlags, ExamNotes, ExamType, ExamTypeValues ) " & _
    "SELECT mInstrumentID, mExamRefNum, mInstExamName, mHISExamCode, mDisplayName, mActive, mHISResultCode, mUnit, mDiff, " & _
    "mDilution, mFactorOP, mFormula, mHiloFlags, mExamNotes, mExamType, mExamTypeValues;"

    qryDisplayProfileExams = "PARAMETERS mProfileID Text ( 255 ), mInstrumentID Text ( 255 ); " & _
    "SELECT string(4 * val(Profiles.parentID) ,' ') + Exams.DisplayName AS [Exam Name], '' AS [Instrument Result], " & _
    "'' AS Result, '' AS [Normal Values], '' AS Status, Exams.InstExamName, Profiles.Dilution,  " & _
    "Profiles.DataFormat, Profiles.ExamID, Profiles.SortNum, Exams.Diffcount, '' AS Units, Exams.FactorOP, " & _
    "'' AS [Unit (SI)], '' AS [Result (SI)], '' AS [Normal Values (SI)], Exams.Dilution AS DilutionSI,  " & _
    "Profiles.Group AS Profile, Exams.Formula, Exams.HiloFlags, Exams.ExamNotes, Exams.ExamType, Exams.ExamTypeValues " & _
    "FROM tbLabProfileExams AS Profiles LEFT JOIN tbCoLabExam AS Exams ON Profiles.ExamID = Exams.examID " & _
    "WHERE Exams.Active = TRUE and Profiles.Profileid = mProfileID and Profiles.InstrumentID = mInstrumentID " & _
    "ORDER BY Profiles.SortNum;"

    qryVSValSelect = "PARAMETERS mRefNum Text ( 255 ); " & _
    "SELECT Patient.*, Exams.*, ProfileExams.*, InstExam.InstExamName, InstExam.HISExamCode, InstExam.HisResultCode " & _
    "FROM tbPatientValues AS Exams, tbPatientMaster AS Patient, tbLabProfileExams AS ProfileExams, tbCoLabExam as InstExam " & _
    "WHERE (((Exams.RefNum)=[mRefNum]) AND ((Patient.Refnum)=[mRefNum]) AND ((ProfileExams.ProfileID)=[Patient].[ProfileID]) AND " & _
    "((ProfileExams.ExamID)=[Exams].[ExamID])  and (InstExam.InstrumentID = ProfileExams.InstrumentID and InstExam.ExamID = ProfileExams.ExamID ) ) order by ProfileExams.Sortnum;"
    
    qryPreviousResult = "PARAMETERS mPatientID Text ( 255 ), mExamID Text ( 255 ), mLastName Text ( 255 ), mFirstName Text ( 255 ), mResultDate text ( 255 ); " & _
    "Select top 10 FORMAT(PatMaster.PrintDate,'MM/dd/yy HH:mm') as [VDATE], PatResult.Displayname as [Exam Name],PatResult.Result as [Result1], PatResult.NValuesDisplay as [NV1], " & _
    "PatResult.ResultSI as [Result2], PatResult.NvaluesSIDisplay as [NV2], IResult.[Status],PatResult.Recheck, PatMaster.PrintDate as [VerifyDate] " & _
    "from (tbPatientMaster as PatMaster " & _
    "left outer join tbPatientValues as PatResult on (PatMaster.RefNum = PatResult.RefNum and len(trim(Patresult.Result)) > 0) ) " & _
    "left outer join tbLabInstrumentResults as IResult on PatMaster.Refnum = IResult.Refnum " & _
    "where patmaster.patientid = mPatientID  and PatResult.ExamID = mExamID and PatMaster.LastName = mLastName and PatMaster.FirstName = mFirstName and IResult.ResultDate < cdate(mResultDate) " & _
    "Order By PatMaster.PrintDate desc "
    
    ctr = Len(DBDatabase)
    Do While Mid(DBDatabase, ctr, 1) <> "\" Or ctr = 0
        ctr = ctr - 1
    Loop
    If ctr > 0 Then
        tempDB = Mid(DBDatabase, 1, ctr) & "LabPatientData.mdb"
        tempDBIntrumentResult = Mid(DBDatabase, 1, ctr) & "InstrumentResults.mdb"
    End If
    If Len(tempDB) > 0 Then
        conTemp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & tempDB & ";User Id=;Password=;Jet OLEDB:Database Password=" & DBPassword & ";"
        conTemp.ConnectionTimeout = 30
        conTemp.CursorLocation = adUseClient
        conTemp.Open
        Set rstTemp = myConnection.Execute("Select top 1 * From tbPatientMaster")
        With rstTemp
            For ctr = 0 To rstTemp.Fields.Count - 1
                If rstTemp.Fields(ctr).Name = "PatLicense" Then hasLicense = True
                If rstTemp.Fields(ctr).Name = "PatLicExpire" Then hasLicExpire = True
                If rstTemp.Fields(ctr).Name = "MTLicense" Then hasPatLicense = True
                If rstTemp.Fields(ctr).Name = "MTLicExpire" Then hasPatLicExpire = True
                If rstTemp.Fields(ctr).Name = "ResultFlag" Then hasResultFlag = True
                If rstTemp.Fields(ctr).Name = "Remarks2" Then hasRemarks2 = True
                If rstTemp.Fields(ctr).Name = "InstrumentResultName" And rstTemp.Fields(ctr).ActualSize < 50 Then ModifyIRName = True
                If rstTemp.Fields(ctr).Name = "Ward" And rstTemp.Fields(ctr).ActualSize < 20 Then ModifyWardSize = True
                If rstTemp.Fields(ctr).Name = "MedTech" And rstTemp.Fields(ctr).ActualSize < 100 Then ModifyMTSize = True
            Next ctr
        End With
        If Not hasLicense Then conTemp.Execute "ALTER TABLE tbPatientMaster ADD COLUMN MTLicense TEXT(25)"
        If Not hasLicExpire Then conTemp.Execute "ALTER TABLE tbPatientMaster ADD COLUMN MTLicExpire TEXT(25)"
        If Not hasPatLicense Then conTemp.Execute "ALTER TABLE tbPatientMaster ADD COLUMN PatLicense TEXT(25)"
        If Not hasPatLicExpire Then conTemp.Execute "ALTER TABLE tbPatientMaster ADD COLUMN PatLicExpire TEXT(25)"
        If Not hasResultFlag Then conTemp.Execute "ALTER TABLE tbPatientMaster ADD COLUMN ResultFlag TEXT(25)"
        If Not hasRemarks2 Then conTemp.Execute "ALTER TABLE tbPatientMaster ADD COLUMN Remarks2 MEMO"
        If ModifyIRName Then conTemp.Execute "ALTER TABLE tbPatientMaster ALTER COLUMN InstrumentResultName TEXT(50)"
        If ModifyWardSize Then
            conTemp.Execute "ALTER TABLE tbPatientMaster ALTER COLUMN Ward TEXT(20)"
            conTemp.Execute "ALTER TABLE tbPatientMaster ALTER COLUMN Room TEXT(20)"
        End If
        If ModifyMTSize Then
            conTemp.Execute "ALTER TABLE tbPatientMaster ALTER COLUMN MedTech TEXT(100)"
        End If
        conTemp.Execute "ALTER TABLE tbPatientMaster ALTER COLUMN Physician TEXT(100)"
        
        Set conTemp = Nothing
        
        conTemp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & tempDBIntrumentResult & ";User Id=;Password=;Jet OLEDB:Database Password=" & DBPassword & ";"
        conTemp.ConnectionTimeout = 30
        conTemp.CursorLocation = adUseClient
        conTemp.Open
        Set rstTemp = conTemp.Execute("Select top 1 * from tbLabInstrumentResults")
        If rstTemp.RecordCount > 0 Then
            If rstTemp.Fields(ctr).Name = "NewPatientName" And rstTemp.Fields(ctr).ActualSize <= 50 Then
                conTemp.Execute "ALTER TABLE tbLabInstrumentResults ALTER COLUMN NewPatientName TEXT(255)"
            End If
        End If
        Set rstTemp = myConnection.Execute("Select top 1 * From tbCoLabExam")
        With rstTemp
            For ctr = 0 To rstTemp.Fields.Count - 1
                If UCase(rstTemp.Fields(ctr).Name) = "EXAMTYPE" Then hasExamType = True
                If UCase(rstTemp.Fields(ctr).Name) = "EXAMTYPEVALUES" Then hasExamTypeValues = True
            Next ctr
        End With
        If Not hasExamType Then myConnection.Execute "ALTER TABLE tbCoLabExam ADD COLUMN ExamType TEXT(2)"
        If Not hasExamTypeValues Then myConnection.Execute "ALTER TABLE tbCoLabExam ADD COLUMN ExamTypevalues TEXT(20)"
        
        Set rstTemp = myConnection.Execute("Select top 1 * From tbLabReportNames")
        With rstTemp
            For ctr = 0 To rstTemp.Fields.Count - 1
                If UCase(rstTemp.Fields(ctr).Name) = "RPTFILE" Then ModifyReportsTable = True
                If UCase(rstTemp.Fields(ctr).Name) = "RPTNAME" And UCase(rstTemp.Fields(ctr).DefinedSize) = 20 Then
                    ModifyReportName = True
                End If
                If UCase(rstTemp.Fields(ctr).Name) = "RPTFILESIZE" Then ModifyReportSize = True
            Next ctr
        End With
        If Not ModifyReportsTable Then
            myConnection.Execute "ALTER TABLE tbLabReportNames ADD COLUMN rptSectionCode TEXT(10)"
            myConnection.Execute "ALTER TABLE tbLabReportNames ADD COLUMN rptFile OLEOBJECT"
            myConnection.Execute "ALTER TABLE tbLabReportNames ADD COLUMN rptModifyDate TEXT(25)"
            myConnection.Execute "ALTER TABLE tbLabReportNames ADD COLUMN rptModifyUser TEXT(25)"
        End If
        If ModifyReportName Then
            myConnection.Execute "ALTER TABLE tbLabReportNames ALTER COLUMN rptName TEXT(50)"
        End If
        If Not ModifyReportSize Then
            myConnection.Execute "ALTER TABLE tbLabReportNames ADD COLUMN rptFileSize LONG"
        Else
            myConnection.Execute "ALTER TABLE tbLabReportNames ALTER COLUMN rptFileSize LONG"
        End If
        
        Set rstTemp = myConnection.Execute("Select top 1 * From tbPhysician")
        With rstTemp
            For ctr = 0 To rstTemp.Fields.Count - 1
                If UCase(rstTemp.Fields(ctr).Name) = "DRNAME" Then ModifyPhysicianSize = True
            Next ctr
        End With
        If Not ModifyPhysicianSize Then
            myConnection.Execute "ALTER TABLE tbPhysician ADD COLUMN drName TEXT(100)"
        Else
            myConnection.Execute "ALTER TABLE tbPhysician ALTER COLUMN drName TEXT(100)"
        End If
    End If
    
    myConnection.Execute "DROP Procedure sp_UpdatePatientData"
    myConnection.Execute "DROP Procedure sp_InsertPatientData"
    myConnection.Execute "DROP Procedure sp_ReportUpdate"
    myConnection.Execute "DROP Procedure sp_ReportInsert"
    myConnection.Execute "DROP Procedure qryInsertExam"
    myConnection.Execute "DROP Procedure qryUpdateExam"
    myConnection.Execute "DROP Procedure sp_displayProfileExams"
    myConnection.Execute "DROP Procedure vs_ValuesSelect"
    myConnection.Execute "DROP Procedure sp_PreviousResult"
    
    myConnection.Execute "CREATE Procedure sp_UpdatePatientData as " & qryUpdate
    myConnection.Execute "CREATE Procedure sp_InsertPatientData as " & qryInsert
    myConnection.Execute "CREATE Procedure sp_ReportUpdate as " & qryReportUpdate
    myConnection.Execute "CREATE Procedure sp_ReportInsert as " & qryReportInsert
    myConnection.Execute "CREATE Procedure qryInsertExam as " & qryInsertExam
    myConnection.Execute "CREATE Procedure qryUpdateExam as " & qryUpdateExam
    myConnection.Execute "CREATE Procedure sp_displayProfileExams as " & qryDisplayProfileExams
    myConnection.Execute "CREATE Procedure vs_ValuesSelect as " & qryVSValSelect
    myConnection.Execute "CREATE Procedure sp_PreviousResult as " & qryPreviousResult
        
    myConnection.Execute "Update tbVersion set Major = '4', Minor = '12', [Date] = #04/13/2020#"
    SaveSetting "Lab-I", "Settings", "UpdatedTable", "1"

    Set rstTemp = Nothing
    Set conTemp = Nothing
    If Len(Trim(ErrLog)) > 0 Then
        MsgBox "Update successful but with errors. Please restart the application." & vbCrLf & "---------------" & vbCrLf & ErrLog, vbOKOnly + vbInformation, "DB Patch"
    Else
        MsgBox "Update successful. Please restart the application.", vbOKOnly + vbInformation, "DB Patch"
    End If
End If
errFound:
    If Err.Number <> 0 Then
        ErrLog = ErrLog & vbCrLf & Err.Number & ": " & Err.Description
        Err.Clear
        Resume Next
    End If
End Property

Public Sub showLabExams()
With frmLabExam
    .Show 1
End With
End Sub

Public Sub showOtherReports()
With frmOtherReports
    frmOtherReports.Show 1
End With
End Sub

Public Sub showFormSettings()
With frmFormSettings
    .Show 1
End With
End Sub

Public Sub showSelectionLists()
With frmSelectionLists
    .Show 1
End With
End Sub

Public Property Let wDBServerType(ByVal vNewValue As String)
    DBServerType = vNewValue
End Property

Public Property Let wDBUserName(ByVal vNewValue As String)
    DBUserName = vNewValue
End Property

Public Property Let wRPTName(ByVal vNewValue As String)
    yRPTName = vNewValue
End Property

Public Property Let wDBPassword(ByVal vNewValue As String)
    DBPassword = vNewValue
End Property

Public Property Let wDBDatabase(ByVal vNewValue As String)
    DBDatabase = vNewValue
End Property

Public Property Let wDBSchema(ByVal vNewValue As String)
    DBSchema = vNewValue
End Property

Public Property Let wDBServer(ByVal vNewValue As String)
    DBServer = vNewValue
End Property

Public Property Let wDBConnectionString(ByVal vNewValue As String)
    DBConnectionString = vNewValue
End Property

Public Property Let wInstrumentID(ByVal vNewValue As String)
    InstrumentID = vNewValue
End Property

Public Property Let wSectionCode(ByVal vNewValue As String)
    SectionCode = vNewValue
End Property

Public Function showSearchCode(mTrans As String, mSearchValue As String, Optional AutoTab As Boolean = True, Optional mTitle As String = "Search Code", Optional showAllRec As Boolean = False, Optional ShowFilterBox As Boolean = False) As String
Dim tempStr As String
Dim rstTemp As ADODB.Recordset
If IsNumeric(mSearchValue) Then
    If showAllRec Then
        Set rstTemp = GetSearchRecordSet(mTrans, mSearchValue)
        With frmSearchCode
            .wShowFilter = ShowFilterBox
            .wRecordSet = rstTemp
            .wSearchValue = mSearchValue
            .lblTitle.Caption = mTitle
            .Show 1
            tempStr = ReturnValueStr2
            LastsearchID = ReturnValueStr
'            If Len(Trim(tempStr)) > 0 Then
'                If AutoTab Then keybd_event VK_TAB, 0, 0, 0
'            End If
        End With
    Else
        tempStr = SearchCode(mTrans, mSearchValue)
        If Len(Trim(tempStr)) = 0 Then
            MsgBox "No Record Found...", vbInformation + vbOKOnly, "Lab-I Search Code"
            tempStr = ""
            highLight
        Else
            If AutoTab Then keybd_event VK_TAB, 0, 0, 0
        End If
        LastsearchID = mSearchValue
    End If
Else
    Set rstTemp = GetSearchRecordSet(mTrans, mSearchValue)
    If rstTemp.RecordCount > 1 Then
        With frmSearchCode
            .wShowFilter = ShowFilterBox
            .wRecordSet = rstTemp
            .wSearchValue = mSearchValue
            .lblTitle.Caption = mTitle
            .Show 1
            tempStr = ReturnValueStr2
            LastsearchID = ReturnValueStr
            If Len(Trim(tempStr)) > 0 Then
                If AutoTab Then keybd_event VK_TAB, 0, 0, 0
            End If
        End With
    ElseIf rstTemp.RecordCount = 1 Then
        If showAllRec Then
            With frmSearchCode
                .wShowFilter = ShowFilterBox
                .wRecordSet = rstTemp
                .wSearchValue = mSearchValue
                .lblTitle.Caption = mTitle
                .Show 1
                tempStr = ReturnValueStr2
                LastsearchID = ReturnValueStr
                If Len(Trim(tempStr)) > 0 Then
                    If AutoTab Then keybd_event VK_TAB, 0, 0, 0
                End If
            End With
        Else
            tempStr = rstTemp.Fields(1).Value
            LastsearchID = rstTemp.Fields(0).Value & ""
            If AutoTab Then keybd_event VK_TAB, 0, 0, 0
        End If
    Else
        MsgBox "No Record Found...", vbInformation + vbOKOnly, "Lab-I Search Code"
        tempStr = ""
        LastsearchID = mSearchValue
        highLight
    End If
End If
showSearchCode = tempStr
LastSearchName = tempStr
End Function

Private Function SearchCode(mTrans As String, mSearchID As String) As String
If DBServerType = "LOCAL DATABASE" Then
    Select Case mTrans
        Case "1"
            If mSearchID = "0" Then
                Set rstData = myConnection.Execute("Select 'Others' as SearchName")
            Else
                Set rstData = myConnection.Execute("Select profilename as SearchName from tbLabProfiles where profileid = '" & mSearchID & "' and InstrumentID = '" & InstrumentID & "'")
            End If
        Case "2"
            Set rstData = myConnection.Execute("Select classname as SearchName from tbLabclassifications where classid = '" & mSearchID & "' and InstrumentID = '" & InstrumentID & "'")
        Case "3"
            Set rstData = myConnection.Execute("Select DisplayName as SearchName from tbcolabexam where examID = '" & mSearchID & "' and Active = true and InstrumentID = '" & InstrumentID & "'")
        Case "4"
            Set rstData = myConnection.Execute("Select rptName as SearchName from tbLabReportNames where rptID = " & mSearchID)
        Case "5"
            Set rstData = myConnection.Execute("Select mtName as SearchName from tbMedTech where mtName = '" & mSearchID & "'")
        Case "6"
            Set rstData = myConnection.Execute("Select drName as SearchName from tbPhysician where drName = '" & mSearchID & "'")
        Case "7"
            Set rstData = myConnection.Execute("Select Pathologist as SearchName from tbLabReportNames where PathologistID = " & mSearchID)
        Case "8"
            Set rstData = myConnection.Execute("Select classname as SearchName from tbLabclassifications where classid = '" & mSearchID & "' and Active = true and InstrumentID = '" & InstrumentID & "'")
        Case "9"
            Set rstData = myConnection.Execute("Select Templates as SearchName from tbLabRemTemplate where Templates = '" & mSearchID & "'")
        Case "10"
            Set rstData = myConnection.Execute("Select PatientID as SearchName from tbPatientMaster where PatientID = '" & mSearchID & "'")
        Case "11"
            Set rstData = myConnection.Execute("Select SelType as SearchName from tbSelectionValues where PatientID = '" & mSearchID & "'")
    End Select
Else
Select Case mTrans
        Case "1"
            If mSearchID = "0" Then
                Set rstData = myConnection.Execute("Select 'Others' as SearchName")
            Else
                Set rstData = myConnection.Execute("Select profilename as SearchName from tbLabProfiles where profileid = '" & mSearchID & "' and InstrumentID = '" & InstrumentID & "'")
            End If
        Case "2"
            Set rstData = myConnection.Execute("Select classname as SearchName from tbLabclassifications where classid = '" & mSearchID & "' and InstrumentID = '" & InstrumentID & "'")
        Case "3"
            Set rstData = myConnection.Execute("Select DisplayName as SearchName from tbcolabexam where examID = '" & mSearchID & "' and Active = 1 and InstrumentID = '" & InstrumentID & "'")
        Case "4"
            Set rstData = myConnection.Execute("Select rptName as SearchName from tbLabReportNames where rptID = " & mSearchID)
        Case "5"
            Set rstData = myConnection.Execute("Select mtName as SearchName from tbMedTech where mtName = '" & mSearchID & "'")
        Case "6"
            Set rstData = myConnection.Execute("Select drName as SearchName from tbPhysician where drName = '" & mSearchID & "'")
        Case "7"
            Set rstData = myConnection.Execute("Select Pathologist as SearchName from tbLabReportNames where PathologistID = " & mSearchID)
        Case "8"
            Set rstData = myConnection.Execute("Select classname as SearchName from tbLabclassifications where classid = '" & mSearchID & "' and Active = 1 and InstrumentID = '" & InstrumentID & "'")
        Case "9"
            Set rstData = myConnection.Execute("Select Templates as SearchName from tbLabRemTemplate where Templates = '" & mSearchID & "'")
        Case "10"
            Set rstData = myConnection.Execute("Select PatientID as SearchName from tbPatientMaster where PatientID = '" & mSearchID & "'")
        Case "11"
            Set rstData = myConnection.Execute("Select SelType as SearchName from tbSelectionValues where PatientID = '" & mSearchID & "'")
End Select
End If
If rstData.RecordCount = 1 Then
    SearchCode = rstData.Fields("SearchName").Value
Else
    SearchCode = ""
End If

End Function

Private Function GetSearchRecordSet(mTrans As String, mSearchID As String) As ADODB.Recordset
Dim tempSearchID As String
tempSearchID = Replace(mSearchID, "'", "''")
If DBServerType = "LOCAL DATABASE" Then
    Select Case mTrans
        Case "1"
            Set GetSearchRecordSet = myConnection.Execute("Select profileid as [Profile ID], profilename as [Profile Name] from tbLabProfiles where profilename like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "' and Status = '1'")
        Case "2"
            Set GetSearchRecordSet = myConnection.Execute("Select classid as [Class ID], classname as [Class Name] from tbLabClassifications where classname like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "'")
        Case "3"
            Set GetSearchRecordSet = myConnection.Execute("Select ExamID as [Exam ID], DisplayName as [Exam Name] from tbcoLabexam where (displayname like '" & tempSearchID & "%' or InstExamName like '" & tempSearchID & "%') and InstrumentID = '" & InstrumentID & "' and Active = True")
        Case "4"
            Set GetSearchRecordSet = myConnection.Execute("Select rptID as [RPT ID], rptName as [RPT Name] from tbLabReportNames where rptName like '" & tempSearchID & "%'")
        Case "5"
            Set GetSearchRecordSet = myConnection.Execute("Select '----'as [ID], mtName as [Medical Technologist Name] from tbMedTech where mtName like '" & tempSearchID & "%'")
        Case "6"
            Set GetSearchRecordSet = myConnection.Execute("Select '----'as [ID], drName as [Physician Name] from tbPhysician where drName like '" & tempSearchID & "%'")
        Case "7"
            Set GetSearchRecordSet = myConnection.Execute("Select PathologistID as [ID], Pathologist as [Pathologist Name] from tbLabPathologists where Pathologist like '" & tempSearchID & "%'")
        Case "8"
            Set GetSearchRecordSet = myConnection.Execute("Select classid as [Class ID], classname as [Class Name] from tbLabClassifications where classname like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "' and Active = True")
        Case "9"
            Set GetSearchRecordSet = myConnection.Execute("Select '----'as [ID], Templates as [Template] from tbLabRemTemplate where Templates like '" & tempSearchID & "%'")
        Case "10"
            Set GetSearchRecordSet = myConnection.Execute("Select PatientID as [ID], Lastname & ', ' & Firstname & ' ' & Middlename as [Patient Name], Age from tbpatientmaster where PatientID like '" & tempSearchID & "%'")
        Case "11"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' as [ID],  Selvalues as [Values] from tbSelectionValues where seltype like '" & tempSearchID & "%'")
    End Select
ElseIf DBServerType = "MYSQL SERVER" Then
    Select Case mTrans
        Case "1"
            Set GetSearchRecordSet = myConnection.Execute("Select profileid ""Profile ID"", profilename ""Profile Name"" from tbLabProfiles where profilename like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "' and Status = '1'")
        Case "2"
            Set GetSearchRecordSet = myConnection.Execute("Select classid ""Class ID"", classname as ""Class Name"" from tbLabClassifications where classname like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "'")
        Case "3"
            Set GetSearchRecordSet = myConnection.Execute("Select ExamID ""Exam ID"", DisplayName ""Exam Name"" from tbcoLabexam where (displayname like '" & tempSearchID & "%' or InstExamName like '" & tempSearchID & "%') and InstrumentID = '" & InstrumentID & "' and Active = True")
        Case "4"
            Set GetSearchRecordSet = myConnection.Execute("Select rptID ""RPT ID"", rptName ""RPT Name"" from tbLabReportNames where rptName like '" & tempSearchID & "%'")
        Case "5"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"", mtName ""Medical Technologist Name"" from tbMedTech where mtName like '" & tempSearchID & "%'")
        Case "6"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"", drName ""Physician Name"" from tbPhysician where drName like '" & tempSearchID & "%'")
        Case "7"
            Set GetSearchRecordSet = myConnection.Execute("Select PathologistID ""ID"", Pathologist ""Pathologist Name"" from tbLabPathologists where Pathologist like '" & tempSearchID & "%'")
        Case "8"
            Set GetSearchRecordSet = myConnection.Execute("Select classid ""Class ID"", classname ""Class Name"" from tbLabClassifications where classname like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "' and Active = True")
        Case "9"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"", Templates ""Template"" from tbLabRemTemplate where Templates like '" & tempSearchID & "%'")
        Case "10"
            Set GetSearchRecordSet = myConnection.Execute("Select PatientID ""ID"", Lastname & ', ' & Firstname & ' ' & Middlename ""Patient Name"", Age from tbpatientmaster where PatientID like '" & tempSearchID & "%'")
        Case "11"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"",  Selvalues as ""Values"" from tbSelectionValues where seltype like '" & tempSearchID & "%'")
    End Select

ElseIf DBServerType = "MSSQL SERVER" Then
    Select Case mTrans
        Case "1"
            Set GetSearchRecordSet = myConnection.Execute("Select profileid ""Profile ID"", profilename ""Profile Name"" from tbLabProfiles where profilename like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "' and Status = '1'")
        Case "2"
            Set GetSearchRecordSet = myConnection.Execute("Select classid ""Class ID"", classname as ""Class Name"" from tbLabClassifications where classname like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "'")
        Case "3"
            Set GetSearchRecordSet = myConnection.Execute("Select ExamID ""Exam ID"", DisplayName ""Exam Name"" from tbcoLabexam where (displayname like '" & tempSearchID & "%' or InstExamName like '" & tempSearchID & "%') and InstrumentID = '" & InstrumentID & "' and Active = 1")
        Case "4"
            Set GetSearchRecordSet = myConnection.Execute("Select rptID ""RPT ID"", rptName ""RPT Name"" from tbLabReportNames where rptName like '" & tempSearchID & "%'")
        Case "5"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"", mtName ""Medical Technologist Name"" from tbMedTech where mtName like '" & tempSearchID & "%'")
        Case "6"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"", drName ""Physician Name"" from tbPhysician where drName like '" & tempSearchID & "%'")
        Case "7"
            Set GetSearchRecordSet = myConnection.Execute("Select PathologistID ""ID"", Pathologist ""Pathologist Name"" from tbLabPathologists where Pathologist like '" & tempSearchID & "%'")
        Case "8"
            Set GetSearchRecordSet = myConnection.Execute("Select classid ""Class ID"", classname ""Class Name"" from tbLabClassifications where classname like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "' and Active = 1")
        Case "9"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"", Templates ""Template"" from tbLabRemTemplate where Templates like '" & tempSearchID & "%'")
        Case "10"
            Set GetSearchRecordSet = myConnection.Execute("Select PatientID ""ID"", Lastname & ', ' & Firstname & ' ' & Middlename ""Patient Name"", Age from tbpatientmaster where PatientID like '" & tempSearchID & "%'")
        Case "11"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"",  Selvalues as ""Values"" from tbSelectionValues where seltype like '" & tempSearchID & "%'")
    End Select

End If

    If DBServerType = "MSSQL SERVER" Then
    Select Case mTrans
        Case "1"
            Set GetSearchRecordSet = myConnection.Execute("Select profileid as [Profile ID], profilename as [Profile Name] from tbLabProfiles where profilename like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "' and Status = '1'")
        Case "2"
            Set GetSearchRecordSet = myConnection.Execute("Select classid as [Class ID], classname as [Class Name] from tbLabClassifications where classname like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "'")
        Case "3"
            Set GetSearchRecordSet = myConnection.Execute("Select ExamID as [Exam ID], DisplayName as [Exam Name] from tbcoLabexam where (displayname like '" & tempSearchID & "%' or InstExamName like '" & tempSearchID & "%') and InstrumentID = '" & InstrumentID & "' and Active = 1")
        Case "4"
            Set GetSearchRecordSet = myConnection.Execute("Select rptID as [RPT ID], rptName as [RPT Name] from tbLabReportNames where rptName like '" & tempSearchID & "%'")
        Case "5"
            Set GetSearchRecordSet = myConnection.Execute("Select '----'as [ID], mtName as [Medical Technologist Name] from tbMedTech where mtName like '" & tempSearchID & "%'")
        Case "6"
            Set GetSearchRecordSet = myConnection.Execute("Select '----'as [ID], drName as [Physician Name] from tbPhysician where drName like '" & tempSearchID & "%'")
        Case "7"
            Set GetSearchRecordSet = myConnection.Execute("Select PathologistID as [ID], Pathologist as [Pathologist Name] from tbLabPathologists where Pathologist like '" & tempSearchID & "%'")
        Case "8"
            Set GetSearchRecordSet = myConnection.Execute("Select classid as [Class ID], classname as [Class Name] from tbLabClassifications where classname like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "' and Active = 1")
        Case "9"
            Set GetSearchRecordSet = myConnection.Execute("Select '----'as [ID], Templates as [Template] from tbLabRemTemplate where Templates like '" & tempSearchID & "%'")
        Case "10"
            Set GetSearchRecordSet = myConnection.Execute("Select PatientID as [ID], Lastname & ', ' & Firstname & ' ' & Middlename as [Patient Name], Age from tbpatientmaster where PatientID like '" & tempSearchID & "%'")
        Case "11"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' as [ID],  Selvalues as [Values] from tbSelectionValues where seltype like '" & tempSearchID & "%'")
    End Select
ElseIf DBServerType = "MYSQL SERVER" Then
    Select Case mTrans
        Case "1"
            Set GetSearchRecordSet = myConnection.Execute("Select profileid ""Profile ID"", profilename ""Profile Name"" from tbLabProfiles where profilename like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "' and Status = '1'")
        Case "2"
            Set GetSearchRecordSet = myConnection.Execute("Select classid ""Class ID"", classname as ""Class Name"" from tbLabClassifications where classname like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "'")
        Case "3"
            Set GetSearchRecordSet = myConnection.Execute("Select ExamID ""Exam ID"", DisplayName ""Exam Name"" from tbcoLabexam where (displayname like '" & tempSearchID & "%' or InstExamName like '" & tempSearchID & "%') and InstrumentID = '" & InstrumentID & "' and Active = True")
        Case "4"
            Set GetSearchRecordSet = myConnection.Execute("Select rptID ""RPT ID"", rptName ""RPT Name"" from tbLabReportNames where rptName like '" & tempSearchID & "%'")
        Case "5"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"", mtName ""Medical Technologist Name"" from tbMedTech where mtName like '" & tempSearchID & "%'")
        Case "6"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"", drName ""Physician Name"" from tbPhysician where drName like '" & tempSearchID & "%'")
        Case "7"
            Set GetSearchRecordSet = myConnection.Execute("Select PathologistID ""ID"", Pathologist ""Pathologist Name"" from tbLabPathologists where Pathologist like '" & tempSearchID & "%'")
        Case "8"
            Set GetSearchRecordSet = myConnection.Execute("Select classid ""Class ID"", classname ""Class Name"" from tbLabClassifications where classname like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "' and Active = True")
        Case "9"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"", Templates ""Template"" from tbLabRemTemplate where Templates like '" & tempSearchID & "%'")
        Case "10"
            Set GetSearchRecordSet = myConnection.Execute("Select PatientID ""ID"", Lastname & ', ' & Firstname & ' ' & Middlename ""Patient Name"", Age from tbpatientmaster where PatientID like '" & tempSearchID & "%'")
        Case "11"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"",  Selvalues as ""Values"" from tbSelectionValues where seltype like '" & tempSearchID & "%'")
    End Select
       
    If DBServerType = "MYSQL SERVER" Then
        Select Case mTrans
        Case "1"
            Set GetSearchRecordSet = myConnection.Execute("Select profileid as [Profile ID], profilename as [Profile Name] from tbLabProfiles where profilename like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "' and Status = '1'")
        Case "2"
            Set GetSearchRecordSet = myConnection.Execute("Select classid as [Class ID], classname as [Class Name] from tbLabClassifications where classname like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "'")
        Case "3"
            Set GetSearchRecordSet = myConnection.Execute("Select ExamID as [Exam ID], DisplayName as [Exam Name] from tbcoLabexam where (displayname like '" & tempSearchID & "%' or InstExamName like '" & tempSearchID & "%') and InstrumentID = '" & InstrumentID & "' and Active = True")
        Case "4"
            Set GetSearchRecordSet = myConnection.Execute("Select rptID as [RPT ID], rptName as [RPT Name] from tbLabReportNames where rptName like '" & tempSearchID & "%'")
        Case "5"
            Set GetSearchRecordSet = myConnection.Execute("Select '----'as [ID], mtName as [Medical Technologist Name] from tbMedTech where mtName like '" & tempSearchID & "%'")
        Case "6"
            Set GetSearchRecordSet = myConnection.Execute("Select '----'as [ID], drName as [Physician Name] from tbPhysician where drName like '" & tempSearchID & "%'")
        Case "7"
            Set GetSearchRecordSet = myConnection.Execute("Select PathologistID as [ID], Pathologist as [Pathologist Name] from tbLabPathologists where Pathologist like '" & tempSearchID & "%'")
        Case "8"
            Set GetSearchRecordSet = myConnection.Execute("Select classid as [Class ID], classname as [Class Name] from tbLabClassifications where classname like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "' and Active = True")
        Case "9"
            Set GetSearchRecordSet = myConnection.Execute("Select '----'as [ID], Templates as [Template] from tbLabRemTemplate where Templates like '" & tempSearchID & "%'")
        Case "10"
            Set GetSearchRecordSet = myConnection.Execute("Select PatientID as [ID], Lastname & ', ' & Firstname & ' ' & Middlename as [Patient Name], Age from tbpatientmaster where PatientID like '" & tempSearchID & "%'")
        Case "11"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' as [ID],  Selvalues as [Values] from tbSelectionValues where seltype like '" & tempSearchID & "%'")
    End Select

  ElseIf DBServerType = "MSSQL SERVER" Then
        Select Case mTrans
        Case "1"
            Set GetSearchRecordSet = myConnection.Execute("Select profileid ""Profile ID"", profilename ""Profile Name"" from tbLabProfiles where profilename like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "' and Status = '1'")
        Case "2"
            Set GetSearchRecordSet = myConnection.Execute("Select classid ""Class ID"", classname as ""Class Name"" from tbLabClassifications where classname like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "'")
        Case "3"
            Set GetSearchRecordSet = myConnection.Execute("Select ExamID ""Exam ID"", DisplayName ""Exam Name"" from tbcoLabexam where (displayname like '" & tempSearchID & "%' or InstExamName like '" & tempSearchID & "%') and InstrumentID = '" & InstrumentID & "' and Active = 1")
        Case "4"
            Set GetSearchRecordSet = myConnection.Execute("Select rptID ""RPT ID"", rptName ""RPT Name"" from tbLabReportNames where rptName like '" & tempSearchID & "%'")
        Case "5"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"", mtName ""Medical Technologist Name"" from tbMedTech where mtName like '" & tempSearchID & "%'")
        Case "6"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"", drName ""Physician Name"" from tbPhysician where drName like '" & tempSearchID & "%'")
        Case "7"
            Set GetSearchRecordSet = myConnection.Execute("Select PathologistID ""ID"", Pathologist ""Pathologist Name"" from tbLabPathologists where Pathologist like '" & tempSearchID & "%'")
        Case "8"
            Set GetSearchRecordSet = myConnection.Execute("Select classid ""Class ID"", classname ""Class Name"" from tbLabClassifications where classname like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "' and Active = True")
        Case "9"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"", Templates ""Template"" from tbLabRemTemplate where Templates like '" & tempSearchID & "%'")
        Case "10"
            Set GetSearchRecordSet = myConnection.Execute("Select PatientID ""ID"", Lastname & ', ' & Firstname & ' ' & Middlename ""Patient Name"", Age from tbpatientmaster where PatientID like '" & tempSearchID & "%'")
        Case "11"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"",  Selvalues as ""Values"" from tbSelectionValues where seltype like '" & tempSearchID & "%'")
End Select

    If DBServerType = "LOCAL DATABASE" Then
        Select Case mTrans
        Case "1"
            Set GetSearchRecordSet = myConnection.Execute("Select profileid ""Profile ID"", profilename ""Profile Name"" from tbLabProfiles where profilename like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "' and Status = '1'")
        Case "2"
            Set GetSearchRecordSet = myConnection.Execute("Select classid ""Class ID"", classname as ""Class Name"" from tbLabClassifications where classname like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "'")
        Case "3"
            Set GetSearchRecordSet = myConnection.Execute("Select ExamID ""Exam ID"", DisplayName ""Exam Name"" from tbcoLabexam where (displayname like '" & tempSearchID & "%' or InstExamName like '" & tempSearchID & "%') and InstrumentID = '" & InstrumentID & "' and Active = True")
        Case "4"
            Set GetSearchRecordSet = myConnection.Execute("Select rptID ""RPT ID"", rptName ""RPT Name"" from tbLabReportNames where rptName like '" & tempSearchID & "%'")
        Case "5"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"", mtName ""Medical Technologist Name"" from tbMedTech where mtName like '" & tempSearchID & "%'")
        Case "6"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"", drName ""Physician Name"" from tbPhysician where drName like '" & tempSearchID & "%'")
        Case "7"
            Set GetSearchRecordSet = myConnection.Execute("Select PathologistID ""ID"", Pathologist ""Pathologist Name"" from tbLabPathologists where Pathologist like '" & tempSearchID & "%'")
        Case "8"
            Set GetSearchRecordSet = myConnection.Execute("Select classid ""Class ID"", classname ""Class Name"" from tbLabClassifications where classname like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "' and Active = True")
        Case "9"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"", Templates ""Template"" from tbLabRemTemplate where Templates like '" & tempSearchID & "%'")
        Case "10"
            Set GetSearchRecordSet = myConnection.Execute("Select PatientID ""ID"", Lastname & ', ' & Firstname & ' ' & Middlename ""Patient Name"", Age from tbpatientmaster where PatientID like '" & tempSearchID & "%'")
        Case "11"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"",  Selvalues as ""Values"" from tbSelectionValues where seltype like '" & tempSearchID & "%'")
    End Select

ElseIf DBServerType = "MSSQL SERVER" Then
    Select Case mTrans
        Case "1"
            Set GetSearchRecordSet = myConnection.Execute("Select profileid ""Profile ID"", profilename ""Profile Name"" from tbLabProfiles where profilename like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "' and Status = '1'")
        Case "2"
            Set GetSearchRecordSet = myConnection.Execute("Select classid ""Class ID"", classname as ""Class Name"" from tbLabClassifications where classname like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "'")
        Case "3"
            Set GetSearchRecordSet = myConnection.Execute("Select ExamID ""Exam ID"", DisplayName ""Exam Name"" from tbcoLabexam where (displayname like '" & tempSearchID & "%' or InstExamName like '" & tempSearchID & "%') and InstrumentID = '" & InstrumentID & "' and Active = 1")
        Case "4"
            Set GetSearchRecordSet = myConnection.Execute("Select rptID ""RPT ID"", rptName ""RPT Name"" from tbLabReportNames where rptName like '" & tempSearchID & "%'")
        Case "5"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"", mtName ""Medical Technologist Name"" from tbMedTech where mtName like '" & tempSearchID & "%'")
        Case "6"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"", drName ""Physician Name"" from tbPhysician where drName like '" & tempSearchID & "%'")
        Case "7"
            Set GetSearchRecordSet = myConnection.Execute("Select PathologistID ""ID"", Pathologist ""Pathologist Name"" from tbLabPathologists where Pathologist like '" & tempSearchID & "%'")
        Case "8"
            Set GetSearchRecordSet = myConnection.Execute("Select classid ""Class ID"", classname ""Class Name"" from tbLabClassifications where classname like '" & tempSearchID & "%' and InstrumentID = '" & InstrumentID & "' and Active = True")
        Case "9"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"", Templates ""Template"" from tbLabRemTemplate where Templates like '" & tempSearchID & "%'")
        Case "10"
            Set GetSearchRecordSet = myConnection.Execute("Select PatientID ""ID"", Lastname & ', ' & Firstname & ' ' & Middlename ""Patient Name"", Age from tbpatientmaster where PatientID like '" & tempSearchID & "%'")
        Case "11"
            Set GetSearchRecordSet = myConnection.Execute("Select '----' ""ID"",  Selvalues as ""Values"" from tbSelectionValues where seltype like '" & tempSearchID & "%'")
    End Select
    End If
End If
End If
End Function

Public Function showAddProfile(Optional mSelectedItems As String = "") As String
With frmAddProfiles
    .wSelectedItems = mSelectedItems
    .Show 1
    showAddProfile = ReturnValueStr
End With
End Function

Public Sub showNormalValues()
With frmNormalValues
    .Show 1
End With
End Sub

Public Sub showHospitalInfo()
With frmHospital
    .Show 1
End With
End Sub

Public Sub showProfileExams()
With frmLabProfileExams
    .Show 1
End With
End Sub

Public Sub showResultPrintPreview(mRefnum As String)
    If OpenMainReport(GetReport(getDefaultTemplate(mRefnum)), mRefnum & "", "") <> "Error" Then
        ShowReportViewer False, "Result Print Preview", False, mRefnum
    End If
End Sub

Public Sub showResultAutoPrint(mProfileID As String, mInstrumentID As String, mAccessionnum As String, mPatientName As String, mResultDate As String)
    If OpenMainReport(App.Path & "\reports\AutoPrint_01.rpt", mProfileID & "", mInstrumentID, mAccessionnum, mPatientName, mResultDate, "") <> "Error" Then
        ShowReportViewer False, "Result Print Preview", False, ""
    End If
End Sub

Public Sub PrintBarCode(mIdNum As String)
    If OpenMainReport(App.Path & "\reports\BarCode.rpt", mIdNum) <> "Error" Then
        ShowReportViewer False, "Bar Code Printing", True
    End If
End Sub


Public Sub ResultAutoPrint(mProfileID As String, mInstrumentID As String, mAccessionnum As String, mPatientName As String, mResultDate As String)
    If OpenMainReport(App.Path & "\reports\AutoPrint_01.rpt", mProfileID & "", mInstrumentID, mAccessionnum, mPatientName, mResultDate, "") <> "Error" Then
        ShowReportViewer False, "Result Print Preview", True, ""
    End If
End Sub


Public Sub showLogbook()
    frmLogbook.Show 1
End Sub

Public Sub showWorklistReport()
    frmWorklistReport.Show 1
End Sub

Public Sub ResultPrint(mRefnum As String)
    If OpenMainReport(GetReport(getDefaultTemplate(mRefnum)), mRefnum & "", "") <> "Error" Then
        ShowReportViewer False, "Result Print Preview", True, mRefnum
    End If
End Sub

Public Sub showReportFormats()
With frmReportFormats
    .Show 1
End With
End Sub

Public Sub showLabProfiles()
With frmLabProfiles
    .Show 1
End With
End Sub

Public Sub showClassifications()
With frmClassifications
    .Show 1
End With
End Sub

Public Sub showPathologist()
With frmPathologist
    .Show 1
End With
End Sub

Public Sub showMedTech()
With frmMedtech
    .Show 1
End With
End Sub

Public Sub showRemarksTemplate()
With frmRemTemplate
    .Show 1
End With
End Sub

Public Function showPIDsearch(myId As String, Optional exDatabase As Boolean = False) As String
 On Error Resume Next
With frmPIDsearch
    .wSearchString = myId
    .wExDataBase = exDatabase
    .Show 1
    showPIDsearch = .wReturnValue
End With
End Function

Public Sub showPhysician()
With frmPhysician
    .Show 1
End With
End Sub

Private Sub Class_Initialize()
On Error GoTo errFound
Dim pluginVSP As String
    TotalDiff = GetSetting("Lab-I", "Settings", "DiffCount", "1")
    EnableSearchID = Val(GetSetting("HostForms_SA", "Settings", "SearchID", "0"))
    EnableVSP = Val(GetSetting("Lab-I", "Versa", "EnablePlugin", "0"))
    If EnableVSP Then
        pluginVSP = Format(GetSetting("Lab-I", "VerSa", "PlugIN", ""), ">")
        Set ObjVSP = CreateObject(Replace(pluginVSP, ".DLL", ".mFunctions"))
    End If
If Val(GetSetting("LAB-I", "OBISettings", "UseExDatabase", "0")) = 1 Then
    If GetSetting("LAB-I", "OBISettings", "ServerType", "") = "REST API" Then
    Else
        If myCon.State = 0 Then
            If GetSetting("LAB-I", "OBISettings", "ServerType", "") = "Informix-CLI 2.5" Then
                myCon.ConnectionString = "Driver={Informix-CLI 2.5 (32 Bit)};Server=" & GetSetting("LAB-I", "OBISettings", "ServerName", "") & ";Database=" & GetSetting("LAB-I", "OBISettings", "Database", "") & ";Uid=" & GetSetting("LAB-I", "OBISettings", "UserName", "") & ";Pwd=" & GetSetting("LAB-I", "OBISettings", "Password", "") & ";"
            ElseIf GetSetting("LAB-I", "OBISettings", "ServerType", "") = "MySQL" Then
                myCon.ConnectionString = "Driver={MySQL ODBC 3.51 Driver};Port=3306;Server=" & GetSetting("LAB-I", "OBISettings", "ServerName", "") & ";Database=" & GetSetting("LAB-I", "OBISettings", "Database", "") & ";User=" & GetSetting("LAB-I", "OBISettings", "UserName", "") & ";Password=" & GetSetting("LAB-I", "OBISettings", "Password", "") & ";Option=3;"
            ElseIf GetSetting("LAB-I", "OBISettings", "ServerType", "") = "MS SQL Server 2000" Then
                myCon.ConnectionString = "Provider=sqloledb;Data Source=" & GetSetting("LAB-I", "OBISettings", "ServerName", "") & ";Initial Catalog=" & GetSetting("LAB-I", "OBISettings", "Database", "") & ";Integrated Security=SSPI;User Id=" & GetSetting("LAB-I", "OBISettings", "UserName", "") & ";Password=" & GetSetting("LAB-I", "OBISettings", "Password", "")
            End If
            myCon.ConnectionTimeout = 30
            myCon.CursorLocation = adUseClient
            myCon.Open
        End If
    End If
End If

errFound:
    If Err.Number <> 0 Then
        Select Case Err.Number
            Case -2147467259
                MsgBox "Unable to connect to server" & vbCrLf & vbCrLf & "Please check database settings.", , "OBI Form"
            Case Else
                MsgBox Err.Description & vbCrLf & "Please Contact your Systems Administrator.", , "OBI Form"
        End Select
    Else
    End If
End Sub

Public Function showInputBox(mTitle As String, Optional mTop As Integer, Optional mLeft As Integer, Optional mValue As String = "", Optional mMaxLength As Integer = 0, Optional mUpCase As Boolean = False, Optional Mask As Boolean = False, Optional MultiLine As Boolean = False) As String
With frmInputBox
    If mTop <> 0 Then .wTop = mTop
    If mLeft <> 0 Then .wLeft = mLeft
    .wTitle = mTitle
    .wValue = mValue
    .wMaxLength = mMaxLength
    .wUpcase = mUpCase
    .wMask = Mask
    .wMultiLine = MultiLine
    .Show 1
    showInputBox = frmSearchVal
End With
End Function

Public Function showPrintCopies() As String
With frmPrintCopies
    .Show 1
    showPrintCopies = ReturnValueStr
End With
End Function

Public Property Get Version() As String
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Property
