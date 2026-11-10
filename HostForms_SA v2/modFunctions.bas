Attribute VB_Name = "modFunctions"
Option Explicit
Public myConnection As New ADODB.Connection
Public Report As New CRAXDRT.Report

Public DBServerType As String
Public DBUserName As String
Public DBPassword As String
Public DBDatabase As String
Public DBSchema As String
Public DBServer As String
Public DBConnectionString As String

Public yRPTName As String

Public InstrumentID As String
Public SectionCode As String

Public rstMedtech As ADODB.Recordset
Public rstPhysician As ADODB.Recordset
Public rstHospital As ADODB.Recordset

Public ObjVSP As Object
Public EnableVSP As Integer

Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
    Public Const LB_FINDSTRING = &H18F

'Sendkeys replacement
Public Const VK_TAB = 9
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2
Public Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public ReturnValueStr As String
Public ReturnValueStr2 As String
Public ReturnValueInt As String
Public TotalDiff As String
Public EnableSearchID As Integer

Public frmSearchVal As String
Public TemplateValue As String

Public myCon As New ADODB.Connection

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" _
(ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Declare Function SHGetPathFromIDList Lib "Shell32.dll" Alias "SHGetPathFromIDListA" _
(ByVal pidl As Long, ByVal pszPath As String) As Long
Public Type SHITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As SHITEMID
End Type
Public Const MAX_PATH As Integer = 260

Public Function PtrObj(ByVal Pointer As Long) As Object
Dim objObject   As Object
    CopyMemory objObject, Pointer, 4&
    Set PtrObj = objObject
    CopyMemory objObject, 0&, 4&
End Function


Public Function OpenMainReport(ByVal strReportFileName As String, ParamArray strParameters()) As String

   Dim crxParameterField As CRAXDRT.ParameterFieldDefinition

   Dim intCtr As Integer
   Dim intTotalParam As Integer
   Dim crxtable As CRAXDRT.DatabaseTable
   Dim crxApplication As New CRAXDRT.Application
   Dim TempTitle As String
   On Error GoTo errFound
   Set Report = Nothing
   TempTitle = ""
   If DBServerType = "LOCAL DATABASE" Then
        If DetectDSN(ODBC_USER_DNS, "Laboratory") Then
            ModifyDSN ODBC_USER_DNS, MICROSOFT_ACCESS, "Laboratory", "Microsoft Access Database", DBDatabase
        Else
            CreateDSN ODBC_USER_DNS, MICROSOFT_ACCESS, "Laboratory", "Microsoft Access Database", DBDatabase
        End If
   Else
        If DetectDSN(ODBC_USER_DNS, "Laboratory") Then
            ModifyDSN ODBC_USER_DNS, MICROSOFT_SQL_SERVER, "Laboratory", "Microsoft SQL SERVER", DBDatabase, DBServer
        Else
            CreateDSN ODBC_USER_DNS, MICROSOFT_SQL_SERVER, "Laboratory", "Microsoft SQL SERVER", DBDatabase, DBServer
        End If
   End If
   
   intTotalParam = UBound(strParameters)
   
   Set Report = crxApplication.OpenReport(strReportFileName, 1)
   For Each crxtable In Report.Database.Tables
      crxtable.SetLogOnInfo "Laboratory", DBDatabase, DBUserName, DBPassword
   Next

   If intTotalParam >= 0 Then
      For Each crxParameterField In Report.ParameterFields
         If intTotalParam >= intCtr Then
            crxParameterField.AddCurrentValue strParameters(intCtr)
            If TempTitle = "" Then
                TempTitle = strParameters(intCtr)
            Else
                TempTitle = TempTitle & "_" & strParameters(intCtr)
            End If
         Else
            Exit For
         End If
         intCtr = intCtr + 1
      Next
      Report.ReportTitle = TempTitle
   End If
errFound:
    If Err.Number = -2147206460 Then
        MsgBox "Result format not found. Please setup a Result format first.", vbOKOnly + vbExclamation, "Result Form"
        OpenMainReport = "Error"
    ElseIf Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbOKOnly + vbExclamation, "Result Form"
        OpenMainReport = "Error"
    Else
        OpenMainReport = ""
    End If
    Set crxParameterField = Nothing
    Set crxtable = Nothing
    Set crxApplication = Nothing
'    Set Report = Nothing
End Function

Public Sub OpenSubReport(ByVal strSubReport As String, ParamArray strParameters())
  Dim crxParameterField As CRAXDRT.ParameterFieldDefinition
   Dim intCtr As Integer
   Dim intTotalParam As Integer
   Dim crxtable As CRAXDRT.DatabaseTable
   Dim crxSubreport As CRAXDRT.Report

   intTotalParam = UBound(strParameters)
   Set crxSubreport = Report.OpenSubReport(strSubReport)

   For Each crxtable In crxSubreport.Database.Tables
      crxtable.SetLogOnInfo "Laboratory", DBDatabase, DBUserName, DBPassword
'     crxtable.Location = DBDatabase
 '    crxtable.SetLogOnInfo crxtable.LogOnServerName, DBDatabase, crxtable.LogOnUserID, DBPassword
   '   crxtable.Location = user.Database + Mid(crxtable.Location, InStr(1, crxtable.Location, "."))
'      crxtable.SetLogOnInfo User.ServerName, User.Database, "sa", User.ServerPassword
   Next

   If intTotalParam >= 0 Then
      For Each crxParameterField In crxSubreport.ParameterFields
         If intTotalParam >= intCtr Then
            crxParameterField.AddCurrentValue strParameters(intCtr)
         Else
            Exit For
         End If
         intCtr = intCtr + 1
      Next
   End If
   Set crxParameterField = Nothing
   Set crxtable = Nothing
   Set crxSubreport = Nothing
End Sub

Public Sub ShowReportViewer(Optional blnShowGroupTree As Boolean = False, Optional strCaption As String = "", Optional blnDirectToPrinter As Boolean = False, Optional mRefnum As String = "", Optional myCopies As Integer = 1, Optional myTag As String = "")
Dim frmView1 As New frmRptViewer
Dim strFormat As Integer
Dim FileExt As String
If myCopies = 0 Then
 Set Report = Nothing
 Exit Sub
End If
If blnDirectToPrinter Then
 Report.Database.Verify
 Report.DisplayProgressDialog = True
 Report.PrintOut False, myCopies
    If strCaption = "Result Print Preview" Then
        If DBServerType = "LOCAL DATABASE" Then
            myConnection.Execute "Update tbLabInstrumentResults set Printed = TRUE where refnum ='" & mRefnum & "'"
            myConnection.Execute "Update tbPatientMaster set Printed = TRUE, PrintDate = #" & Format(Now, "mm/dd/yyyy HH:mm") & "# where refnum ='" & mRefnum & "'"
        Else
            myConnection.Execute "Update tbLabInstrumentResults set Printed = 1 where refnum ='" & mRefnum & "'"
            myConnection.Execute "Update tbPatientMaster set Printed = 1, PrintDate = '" & Format(Now, "mm/dd/yyyy HH:mm") & "' where refnum ='" & mRefnum & "'"
        End If
    End If

    If Val(GetSetting("Lab-I", "VerSa", "EnableExport", "0")) = 1 Then
        Select Case GetSetting("Lab-I", "VerSa", "ExportType", "")
            Case "Comma Separated Values (CSV)"
                strFormat = 5
                FileExt = ".csv"
            Case "Excel (XLS)"
                strFormat = 29
                FileExt = ".xls"
            Case "Rich Text Format (RTF)"
                strFormat = 4
                FileExt = ".rtf"
            Case "HTML"
                strFormat = 32
                FileExt = ".html"
            Case "Word for Windows (DOC)"
                strFormat = 14
                FileExt = ".doc"
            Case "Text File"
                strFormat = 8
                FileExt = ".txt"
            Case "PDF File (PDF)"
                strFormat = 31
                FileExt = ".pdf"
        End Select
        Report.ExportOptions.DestinationType = crEDTDiskFile
        Report.ExportOptions.FormatType = strFormat
        If mRefnum = "" Then
            Report.ExportOptions.ApplicationFileName = strCaption & "_" & Format(Now, "yyyymmdd") & "." & FileExt
        Else
            Report.ExportOptions.ApplicationFileName = strCaption & "_" & mRefnum & "." & FileExt
        End If
        Report.ExportOptions.DiskFileName = GetSetting("Lab-I", "VerSa", "ExportLocation", "") & myTag & FileExt
        Report.Export False
    End If
' Set Report = Nothing
Else
   With frmRptViewer
      .Caption = strCaption
      .directPrint = blnDirectToPrinter
      .wRefNum = mRefnum
      .Show 1
   End With
End If
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

Public Sub setFormColors(myForm As Form)
Dim ctr As Integer
On Error Resume Next
    myForm.fraTitle.BackColor = GetSetting("Lab-I", "FormColors", "TitleBar", &HA66B2B)
    myForm.lblTitle.ForeColor = GetSetting("Lab-I", "FormColors", "TitleText", &HFFFFFF)
    For ctr = 0 To myForm.fraBody.UBound
        myForm.fraBody(ctr).BackColor = GetSetting("Lab-I", "FormColors", "BodyFrame", &H6B4B01)
    Next ctr
    myForm.lblStatus.ForeColor = GetSetting("Lab-I", "FormColors", "BodyText", &HFFFFFF)
    For ctr = 0 To myForm.lblText.UBound
        myForm.lblText(ctr).ForeColor = GetSetting("Lab-I", "FormColors", "BodyText", &HFFFFFF)
    Next ctr
    myForm.msfgResult.BackColorFixed = GetSetting("Lab-I", "FormColors", "TitleBar", &HA66B2B)
    myForm.msfgResult.ForeColorFixed = GetSetting("Lab-I", "FormColors", "TitleText", &HFFFFFF)
End Sub

Public Function getDefaultTemplate(myRefNum As String) As String
Dim rstData As ADODB.Recordset
Dim rstProfile As ADODB.Recordset
Dim rstDefaultReport As ADODB.Recordset
On Error GoTo errFound
Set rstProfile = myConnection.Execute("Select ProfileID from tbPatientMaster where InstrumentID = '" & InstrumentID & "' and RefNum = '" & myRefNum & "'")
Set rstData = myConnection.Execute("Select template from tblabprofiles where instrumentid = '" & InstrumentID & "' and profileid = '" & rstProfile.Fields("ProfileID").Value & "'")
Set rstDefaultReport = myConnection.Execute("Select rptName from tbLabReportNames where rptSectionCode = '" & SectionCode & "' and rptdefault = TRUE")
If rstDefaultReport.RecordCount > 0 Then
    yRPTName = rstDefaultReport.Fields("rptName").Value & ""
Else
    yRPTName = ""
End If
If rstData.RecordCount > 0 Then
    If Len(Trim(rstData.Fields("Template").Value & "")) > 0 Then
        getDefaultTemplate = rstData.Fields("Template").Value & ""
    Else
        getDefaultTemplate = yRPTName
    End If
Else
    getDefaultTemplate = yRPTName
End If
errFound:
    If Err.Number <> 0 Then
        getDefaultTemplate = yRPTName
        Exit Function
    End If
Set rstProfile = Nothing
Set rstData = Nothing
Set rstDefaultReport = Nothing
End Function

Public Function GetNumDay(testYear As Integer, testMonth As Integer) As Integer
    GetNumDay = Day(DateSerial(testYear, testMonth + 1, 0))
End Function

Public Function GenerateResultID(myStringValue As String) As String
Dim strLen As Integer
Dim ctr As Integer
Dim strTemp As String
strLen = Len(myStringValue)
strTemp = ""
For ctr = 1 To strLen
    strTemp = strTemp & Hex$(Asc(Mid(myStringValue, ctr, 1)))
Next ctr
GenerateResultID = strTemp
End Function

Public Function GetReport(mFileName As String) As String
Dim mFileStream As New ADODB.Stream
Dim rstTemp As ADODB.Recordset
Dim mReportFile As String
Dim tempFilePath As String
Set rstTemp = myConnection.Execute("Select rptFile, rptFileSize from tbLabReportnames where rptSectionCode = '" & SectionCode & "' and rptName = '" & mFileName & "'")
If rstTemp.RecordCount > 0 Then
    If mFileStream.State = adStateOpen Then mFileStream.Close
    With mFileStream
        .Type = adTypeBinary
        .Open
        .Write rstTemp.Fields("rptFile").Value
        tempFilePath = IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\" & mFileName
        .SaveToFile tempFilePath, adSaveCreateOverWrite
        GetReport = tempFilePath
    End With
End If
Set rstTemp = Nothing
End Function

Public Sub RemReport(mFileName As String)
    Kill mFileName
End Sub

Public Function ShowExportReport(mFileName As String) As String
With frmRptExport
    ReturnValueStr = ""
    .wFilename = mFileName
    .Show 1
    ShowExportReport = ReturnValueStr
    ReturnValueStr = ""
End With
End Function

Public Function fGetSpecialFolder(mForm As Form, CSIDL As Long) As String
    Dim sPath As String
    Dim iDL As ITEMIDLIST
    '
    ' Retrieve info about system folders such as the "Recent Documents" folder.
    ' Info is stored in the IDL structure.
    '
    fGetSpecialFolder = ""
    If SHGetSpecialFolderLocation(mForm.hWnd, CSIDL, iDL) = 0 Then
        '
        ' Get the path from the ID list, and return the folder.
        '
        sPath = Space$(MAX_PATH)
        If SHGetPathFromIDList(ByVal iDL.mkid.cb, ByVal sPath) Then
            fGetSpecialFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1) & ""
        End If
    End If
End Function
