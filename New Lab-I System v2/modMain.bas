Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const myKey = "3BlindMice"
Public InstrumentID As String
Public MessageFormat As String
Public SectionCode As String
Public PluginVersion As String
Public DefaultProfileID As String
Public ShowActivation As Boolean
Public MsgDef As Object
Public ObjDTP As Object
Public EnableDTP As Integer

'Sendkeys replacement
Public Const VK_TAB = 9
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2
Public Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


Public Declare Function FindWindow Lib "user32" Alias _
"FindWindowA" (ByVal lpClassName As String, ByVal _
lpWindowName As String) As Long

Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
    Public Const LB_FINDSTRING = &H18F
    Public byKey As Boolean ' is list1_click Event activated by a key press ?
    Public byCode As Boolean ' is text1_change Event activated by code ?

Private Declare Sub GlobalMemoryStatusEx Lib "kernel32" (lpBuffer As MEMORYSTATUSEX)

Private Type INT64
   LoPart As Long
   HiPart As Long
End Type

Private Type MEMORYSTATUSEX
   dwLength As Long
   dwMemoryLoad As Long
   ulTotalPhys As INT64
   ulAvailPhys As INT64
   ulTotalPageFile As INT64
   ulAvailPageFile As INT64
   ulTotalVirtual As INT64
   ulAvailVirtual As INT64
   ulAvailExtendedVirtual As INT64
End Type

Public Const frmRcvWidth = 4820
Public Const frmRcvHeight = 3885
Public Blowfish As New clsBlowfish

Public HostSA As Object
Public mComSet As Object

Public DBServerType As String
Public DBUserName As String
Public DBPassword As String
Public DBDatabase As String
Public DBSchema As String
Public DBServer As String
Public DBConnectionString As String
Public initScreenSize As Integer

Public RPTName As String
Public DetectCon As Integer
Public ProcessDateOnly As Integer

Public isConnected As Boolean
Public isResultVisible As Boolean
Public isResultNumeric As Boolean
Public saveStringResult As Boolean
Public Autoprocess As Boolean
Public Autoprint As Boolean
Public MiddlewareMode As Boolean     'Lab-I\Middleware\Enable: forward results to LIS/HIS only (see modOutbox)

Public myConnection As New ADODB.Connection
Public isDBconnected As Boolean
Public rstResult As New ADODB.Recordset
Public TimeOut As Boolean

Public ReturnValue As String
Public ExitProgram As Boolean

Public ActivationDate As String
Public ExpirationDate As String
Public ResultEditingMethod As String

Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" _
    (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" _
    (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" _
    (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" _
    (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal hSrcWidth As Long, _
    ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Const COLORONCOLOR = 3

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public frmSearchVal As String
'Public Const Exams As String = "HGB,HCT,RBC,WBC,MCV,MCH,MCHC,PLT,RDW,LY%,MO%,NE%,EO%,BA%"
Public arrExams As Variant


Public Enum CommStatus
    cs_ReceiveOnly = 1
    cs_SendReceive = 2
    cs_SendID = 3
End Enum

Public myCommStat As CommStatus

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Declare Function GetVersionExA Lib "kernel32" _
               (lpVersionInformation As OSVERSIONINFO) As Integer
               
Public Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

Public Const VER_PLATFORM_WIN32s = 0        ' Win32s on Windows 3.1
Public Const VER_PLATFORM_WIN32_WINDOWS = 1 ' Windows 95, Windows 98, or Windows Me
Public Const VER_PLATFORM_WIN32_NT = 2      ' Windows NT, Windows 2000, Windows XP, or Windows Server

Public Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hWnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long
Public Const MAX_PATH = 260

Public Const charSet As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%"

Public fsoFiles As Object
Public strBackupPath As String
Public iSockets As Integer
Public LoopBack As Boolean
Public SearchReset As Boolean

Public dFilename As String

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

Private Declare Function GetVolumeInformation _
    Lib "kernel32" Alias "GetVolumeInformationA" _
    (ByVal lpRootPathName As String, _
    ByVal pVolumeNameBuffer As String, _
    ByVal nVolumeNameSize As Long, _
    lpVolumeSerialNumber As Long, _
    lpMaximumComponentLength As Long, _
    lpFileSystemFlags As Long, _
    ByVal lpFileSystemNameBuffer As String, _
    ByVal nFileSystemNameSize As Long) As Long
    
Public RegisterOK As Boolean
Public BackupSched As String

Private Declare Function GetRTTAndHopCount _
    Lib "iphlpapi.dll" _
   (ByVal lDestIPAddr As Long, _
    ByRef lHopCount As Long, _
    ByVal lMaxHops As Long, _
    ByRef lRTT As Long) As Long
        
Private Declare Function inet_addr _
    Lib "wsock32.dll" _
   (ByVal cp As String) As Long


Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As _
    Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal lpReserved As Long, lpType As Long, lpData As Any, _
    lpcbData As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As _
    Any, Source As Any, ByVal numBytes As Long)
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
    lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
    lpData As Any, lpcbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
    ByVal cbName As Long) As Long
    
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

Const KEY_READ = &H20019  ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
                          ' KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not
                          ' SYNCHRONIZE))
Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_MULTI_SZ = 7
Const ERROR_MORE_DATA = 234

 
Private Function CLargeInt(Lo As Long, hi As Long) As Double
   Dim dblLo As Double
   Dim dblHi As Double
   If Lo < 0 Then
      dblLo = 2 ^ 32 + Lo
   Else
      dblLo = Lo
   End If
   If hi < 0 Then
      dblHi = 2 ^ 32 + hi
   Else
      dblHi = hi
   End If
   CLargeInt = dblLo + dblHi * 2 ^ 32

End Function


Private Function NumberInKB(ByVal vNumber As Currency) As String
   Dim strReturn As String
   Select Case vNumber
      Case Is < 1024 ^ 1
         strReturn = CStr(vNumber) & " bytes"
      Case Is < 1024 ^ 2
         strReturn = CStr(Round(vNumber / 1024, 1)) & " KB"
      Case Is < 1024 ^ 3
         strReturn = CStr(Round(vNumber / 1024 ^ 2, 2)) & " MB"
      Case Is < 1024 ^ 4
         strReturn = CStr(Round(vNumber / 1024 ^ 3, 2)) & " GB"
   End Select
   NumberInKB = strReturn
End Function

Public Function Ping(prmIPaddr As String) As Boolean
Dim IPaddr As Long, HopsCount As Long, RTT As Long
Dim MaxHops As Long
    Const SUCCESS = 1
    MaxHops = 20
    IPaddr = inet_addr(prmIPaddr)
    Ping = (GetRTTAndHopCount(IPaddr, HopsCount, MaxHops, RTT) = SUCCESS)
End Function

Private Sub Main()
Dim rstData As New ADODB.Recordset
Dim tempstr As String
Dim tempStr2 As String
Dim mExpirationDate As Date
Dim totaldays As Integer
Dim Count As Integer
Dim saveCount As Integer
Dim computedCount As Integer
Dim Thwnd As Long
Dim arrCommand() As String
Dim ctr As Integer
Dim DBLocation As String
Set fsoFiles = CreateObject("Scripting.FileSystemObject")
RegisterOK = False
arrCommand = Split(Replace(Command, "  ", " "), " ")
For ctr = 0 To UBound(arrCommand)
    Select Case UCase(Trim(arrCommand(ctr)))
        Case "-RESETEXPIRATION"
            If MsgBox("Are you sure you want to reset expiration date?", vbYesNo + vbQuestion + vbDefaultButton2, "Reset Expiration Date") = vbYes Then
                DeleteSetting "LAB-i", "Program", "Settings"
                DeleteSetting "LAB-i", "Program", "Install"
                DeleteSetting "LAB-i", "Program", "Path"
                DeleteSetting "LAB-i", "Settings", "VistaFolder"
                MsgBox "Reset Expiration Done. Press OK to close application.", vbOKOnly + vbInformation, "Reset Expiration"
                End
            End If
        Case "-REPATCHDB"
            SaveSetting "Lab-I", "Settings", "UpdatedTable", "0"
            MsgBox "Database Patch has been requested. Press OK to close application.", vbOKOnly + vbInformation, "DB Patch"
            End
        Case "-SAFEMODE"
    End Select
Next ctr

DefaultProfileID = GetSetting("Lab-I", "Settings", "ProfileID", "1")
Autoprocess = Val(GetSetting("Lab-I", "Settings", "Autoprocess", "0"))
Autoprint = Val(GetSetting("Lab-I", "Settings", "Autoprint", "0"))
InitPlugin
iSockets = 0
LoopBack = False
ExitProgram = False
tempstr = DateAdd("yyyy", 1, Now)  'Blowfish.DecryptString(GetSetting("LAB-i", "Program", "Settings", "NoExpire"), myKey, True)
tempStr2 = Now 'tempStr 'Blowfish.DecryptString(GetSetting("LAB-i", "Program", "Install", Blowfish.EncryptString(Format(Now, "mm/dd/yyyy"), myKey, True)), myKey, True)
ActivationDate = Now 'tempStr2
ExpirationDate = DateAdd("yyyy", 1, Now) 'tempStr
GetDBSettings
DBLocation = Replace(Format(DBDatabase, ">"), "LABORATORY.MDB", "")
CopyToVistaFolder
'If Val(GetSetting("Lab-I", "Settings", "VistaFolder", "0")) = 0 Then
'    CopyToVistaFolder GetSpecialFolderPath(&H5)
''    SaveSetting "Lab-I", "Settings", "VistaFolder", "1"
'End If
If IsRegistered Then
    If DateDiff("d", Now, Format(GetSetting("Lab-I", "Registration", "Expiration", ""), "mm/dd/yyyy")) = 30 Or _
        DateDiff("d", Now, Format(GetSetting("Lab-I", "Registration", "Expiration", ""), "mm/dd/yyyy")) = 15 Or _
        (DateDiff("d", Now, Format(GetSetting("Lab-I", "Registration", "Expiration", ""), "mm/dd/yyyy")) <= 7 And _
        DateDiff("d", Now, Format(GetSetting("Lab-I", "Registration", "Expiration", ""), "mm/dd/yyyy")) > 0) Then
            MsgBox "The Laboratory Instrument Interface System will expire on " & Format(GetSetting("Lab-I", "Registration", "Expiration", ""), "dddd, MMMM dd, yyyy") & vbCrLf & vbCrLf & _
                "Please contact Marsman Drysdale Medical Products, Inc. as soon as possible." & vbCrLf & vbCrLf & "Internet connection is required to retrieve the license.", vbOKOnly + vbInformation, App.ProductName
        ElseIf DateDiff("d", Now, Format(GetSetting("Lab-I", "Registration", "Expiration", ""), "mm/dd/yyyy")) = 0 Then
            MsgBox "The Laboratory Instrument Interface System has expired." & vbCrLf & vbCrLf & _
                "Please contact Marsman Drysdale Medical Products, Inc. as soon as possible." & vbCrLf & vbCrLf & "Internet connection is required to retrieve the license.", vbOKOnly + vbInformation, App.ProductName
            frmLicenseManager.Show 1
            If Not IsRegistered Then Exit Sub
    End If
Else
    frmLicenseManager.Show 1
    If Not IsRegistered Then Exit Sub
End If

If IsDate(tempstr) And IsDate(tempStr2) Then
    mExpirationDate = CDate(tempstr)
    totaldays = DateDiff("d", CDate(tempStr2), CDate(tempstr))
    Count = DateDiff("d", Now, mExpirationDate)
    computedCount = (totaldays - Count)
    saveCount = Blowfish.DecryptString(GetSetting("LAB-i", "Program", "Path", Blowfish.EncryptString(Str(computedCount), myKey, True)), myKey, True)
    If Date >= mExpirationDate Or Val(computedCount) < Val(saveCount) Then
        On Error Resume Next
        If Date >= mExpirationDate Then DeleteSetting "LAB-i", "Program", "Path"
        If MsgBox("I'm sorry, The Interface System has reached its expiration date." & vbCrLf & "Please contact Marsman Drysdale Medical Products Inc. as soon as possible." & vbCrLf & _
            "Do you want to enter the activation code now?", vbYesNo + vbDefaultButton2 + vbQuestion, App.ProductName) = vbYes Then
            frmActivation.Show 1
        Else
            Exit Sub
        End If
        End
    Else
        If DateDiff("d", Now, mExpirationDate) = 30 Or DateDiff("d", Now, mExpirationDate) = 20 _
        Or DateDiff("d", Now, mExpirationDate) = 15 Or DateDiff("d", Now, mExpirationDate) = 10 _
        Or DateDiff("d", Now, mExpirationDate) = 7 Or DateDiff("d", Now, mExpirationDate) = 6 _
        Or DateDiff("d", Now, mExpirationDate) = 5 Or DateDiff("d", Now, mExpirationDate) = 4 _
        Or DateDiff("d", Now, mExpirationDate) = 3 Or DateDiff("d", Now, mExpirationDate) = 2 _
        Or DateDiff("d", Now, mExpirationDate) = 1 Then
            MsgBox "The Interface System will expire on " & Format(mExpirationDate, "mmmm dd, yyyy") & ". Please contact Marsman Drysdale Medical Products as soon as possible.", vbOKOnly + vbCritical, App.ProductName
            ShowActivation = True
        End If
    End If
    If Date >= mExpirationDate Then
        SaveSetting "LAB-i", "Program", "Path", Blowfish.EncryptString(Str(computedCount), myKey, True)
        ShowActivation = False
    End If
    isDBconnected = False
    isConnected = False
    isResultVisible = False
    TimeOut = False
    initScreenSize = Screen.Width
    OpenConnection
    If Val(GetSetting("HostForms_SA", "Settings", "SaveStringResult", "0")) = 1 Then saveStringResult = True Else saveStringResult = False
    If DBServerType = "LOCAL DATABASE" Then
        If isDBconnected Then
            VerifyRequestDB Replace(Format(DBDatabase, ">"), "LABORATORY.MDB", "")
            FixLink DBDatabase
            setResultReport
            isResultNumeric = checkResultString
        End If
    End If
    If App.PrevInstance = False Then
        mdiMain.Show
        mdiMain.mnuConnServ.Visible = Not isDBconnected
    Else
        MsgBox "Lab-I Interface System is already running.", vbOKOnly + vbInformation, App.ProductName
        Exit Sub
    End If
Else
    MsgBox "Program Settings has not been set. Please run the program settings first or Enter Activation Code.", , App.ProductName
    SaveSetting "LAB-i", "LOCAL", "DATABASE", Blowfish.EncryptString(App.Path & "\Data\Laboratory.mdb", myKey, True)
    SaveSetting "LAB-i", "LOCAL", "DATABASEPASSWORD", Blowfish.EncryptString("instantiation", myKey, True)
    frmActivation.Show 1
    End
End If
End Sub

'Switches between Main mode (normal Lab-I) and Middleware mode.
'Going to Middleware saves the Main mode Autoprocess/Autoprint values first;
'going back to Main restores them, so the site is exactly as it was.
Public Sub SetMiddlewareMode(ByVal TurnOn As Boolean)
    If TurnOn Then
        If Val(GetSetting("Lab-I", "Middleware", "Enable", "0")) <> 1 Then
            SaveSetting "Lab-I", "Middleware", "SavedAutoprocess", GetSetting("Lab-I", "Settings", "Autoprocess", "0")
            SaveSetting "Lab-I", "Middleware", "SavedAutoprint", GetSetting("Lab-I", "Settings", "Autoprint", "0")
        End If
        SaveSetting "Lab-I", "Middleware", "Enable", "1"
    Else
        SaveSetting "Lab-I", "Middleware", "Enable", "0"
        SaveSetting "Lab-I", "Settings", "Autoprocess", GetSetting("Lab-I", "Middleware", "SavedAutoprocess", GetSetting("Lab-I", "Settings", "Autoprocess", "0"))
        SaveSetting "Lab-I", "Settings", "Autoprint", GetSetting("Lab-I", "Middleware", "SavedAutoprint", GetSetting("Lab-I", "Settings", "Autoprint", "0"))
        On Error Resume Next
        DeleteSetting "Lab-I", "Middleware", "SavedAutoprocess"
        DeleteSetting "Lab-I", "Middleware", "SavedAutoprint"
        On Error GoTo 0
    End If
    'The new mode takes effect when Interface.exe restarts (InitPlugin applies it).
End Sub

Public Sub InitPlugin()
Dim EnablePlugin As Integer
Dim PluginName As String
Dim PluginDTP As String
Dim PluginVSP As String
EnablePlugin = Val(GetSetting("Lab-I", "Settings", "EnablePlugin", "0"))
EnableDTP = Val(GetSetting("Lab-I", "DaTran", "EnablePlugin", "0"))
DetectCon = Val(GetSetting("Lab-I", "Settings", "DetectConnection", "1"))
ProcessDateOnly = Val(GetSetting("Lab-I", "Settings", "DateOnly", "0"))
ResultEditingMethod = GetSetting("Lab-I", "Settings", "Method", "B")
Autoprocess = Val(GetSetting("Lab-I", "Settings", "Autoprocess", "0"))
Autoprint = Val(GetSetting("Lab-I", "Settings", "Autoprint", "0"))
MiddlewareMode = (Val(GetSetting("Lab-I", "Middleware", "Enable", "0")) = 1)
If MiddlewareMode Then
    'Middleware only: every result is processed and forwarded, nothing is printed
    Autoprocess = True
    Autoprint = False
    SaveSetting "Lab-I", "Settings", "Autoprocess", "1"
    SaveSetting "Lab-I", "Settings", "Autoprint", "0"
End If
On Error GoTo errINST
If EnablePlugin = 1 Then
    PluginName = Format(GetSetting("Lab-I", "Settings", "PlugIN", ""), ">")
    Set MsgDef = CreateObject(Replace(PluginName, ".DLL", ".MessageDef"))
    InstrumentID = MsgDef.InstrumentID
    MessageFormat = MsgDef.MessageFormat
    SectionCode = MsgDef.SectionCode
    PluginVersion = MsgDef.Version
Else
    InstrumentID = "HMX"
    MessageFormat = "STKS"
    SectionCode = "HEMA"
    PluginVersion = "1.0.0.0"
End If
If EnableDTP Then
    PluginDTP = Format(GetSetting("Lab-I", "DaTran", "PlugIN", ""), ">")
    Set ObjDTP = CreateObject(Replace(PluginDTP, ".DLL", ".MessageDef"))
End If
errINST:
    If Err.Number = 429 Then
        EnablePlugin = False
        EnableDTP = False
        Resume Next
    End If
End Sub

Public Sub GetDBSettings()
Dim tempstr As String
tempstr = GetSetting("lab-i", "DBSettings", "ServerType", Blowfish.EncryptString("LOCAL DATABASE", myKey, True))
If Len(Trim(tempstr)) > 0 Then
    DBServerType = Blowfish.DecryptString(tempstr, myKey, True)
Else
    DBServerType = "LOCAL DATABASE"
End If
Select Case DBServerType
    Case "LOCAL DATABASE"
        tempstr = GetSetting("lab-i", "LOCAL", "UserName", "")
        If Len(Trim(tempstr)) > 0 Then
            DBUserName = Blowfish.DecryptString(tempstr, myKey, True)
        Else
            DBUserName = ""
        End If
        '~~~~~~~~~~~~~~~~~~
        tempstr = GetSetting("lab-i", "LOCAL", "DATABASE", "")
        If Len(Trim(tempstr)) = 0 Then
            SaveSetting "LAB-i", "LOCAL", "DATABASE", Blowfish.EncryptString(App.Path & "\AppData\Data\Laboratory.mdb", myKey, True)
            SaveSetting "LAB-i", "LOCAL", "DATABASEPASSWORD", Blowfish.EncryptString("instantiation", myKey, True)
        End If
        '~~~~~~~~~~~~~~~~~~
        tempstr = GetSetting("lab-i", "LOCAL", "DATABASEPASSWORD", "")
        If Len(Trim(tempstr)) > 0 Then
            DBPassword = Blowfish.DecryptString(tempstr, myKey, True)
        Else
            DBPassword = ""
        End If
        tempstr = GetSetting("lab-i", "LOCAL", "DATABASE", "")
        If Len(Trim(tempstr)) > 0 Then
            DBDatabase = Blowfish.DecryptString(tempstr, myKey, True)
        Else
            DBDatabase = ""
        End If
        DBServer = "MS Access Database"
        DBConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBDatabase & ";User Id=" & DBUserName & ";Password=;" & "Jet OLEDB:Database Password=" & DBPassword & ";"
    Case "MSSQL SERVER"
        tempstr = GetSetting("lab-i", "MSSQL", "UserName", "")
        If Len(Trim(tempstr)) > 0 Then
            DBUserName = Blowfish.DecryptString(tempstr, myKey, True)
        Else
            DBUserName = ""
        End If
        tempstr = GetSetting("lab-i", "MSSQL", "Password", "")
        If Len(Trim(tempstr)) > 0 Then
            DBPassword = Blowfish.DecryptString(tempstr, myKey, True)
        Else
            DBPassword = ""
        End If
        tempstr = GetSetting("lab-i", "MSSQL", "Database", "")
        If Len(Trim(tempstr)) > 0 Then
            DBDatabase = Blowfish.DecryptString(tempstr, myKey, True)
        Else
            DBDatabase = ""
        End If
        tempstr = GetSetting("lab-i", "MSSQL", "Server", Blowfish.EncryptString("(LOCAL)", myKey, True))
        If Len(Trim(tempstr)) > 0 Then
            DBServer = Blowfish.DecryptString(tempstr, myKey, True)
        Else
            DBServer = ""
        End If
        'DBConnectionString = "Provider=sqloledb;Data Source=" & DBServer & ";Initial Catalog=" & DBDatabase & ";Integrated Security=SSPI;User Id=" & DBUserName & ";Password=" & DBPassword
        DBConnectionString = "Provider=sqloledb;Data Source=" & DBServer & ";Initial Catalog=" & DBDatabase & ";User Id=" & DBUserName & ";Password=" & DBPassword
    Case "MYSQL SERVER"
        tempstr = GetSetting("lab-i", "MYSQL", "UserName", "")
        If Len(Trim(tempstr)) > 0 Then
            DBUserName = Blowfish.DecryptString(tempstr, myKey, True)
        Else
            DBUserName = ""
        End If
        tempstr = GetSetting("lab-i", "MYSQL", "Password", "")
        If Len(Trim(tempstr)) > 0 Then
            DBPassword = Blowfish.DecryptString(tempstr, myKey, True)
        Else
            DBPassword = ""
        End If
        tempstr = GetSetting("lab-i", "MYSQL", "Database", "")
        If Len(Trim(tempstr)) > 0 Then
            DBDatabase = Blowfish.DecryptString(tempstr, myKey, True)
        Else
            DBDatabase = ""
        End If
        tempstr = GetSetting("lab-i", "MYSQL", "Server", Blowfish.EncryptString("(LOCAL)", myKey, True))
        If Len(Trim(tempstr)) > 0 Then
            DBServer = Blowfish.DecryptString(tempstr, myKey, True)
        Else
            DBServer = ""
        End If
        DBConnectionString = "Driver={MySQL ODBC 3.51 Driver};Port=3306;Server=" & DBServer & ";Database=" & DBDatabase & ";User=" & DBUserName & ";Password=" & DBPassword & ";Option=3;"
        '"Provider=sqloledb;Data Source=" & DBServer & ";Initial Catalog=" & DBDatabase & ";Integrated Security=SSPI;User Id=" & DBUserName & ";Password=" & DBPassword
    Case "ORACLE"
        tempstr = GetSetting("lab-i", "ORACLE", "UserName", "")
        If Len(Trim(tempstr)) > 0 Then
            DBUserName = Blowfish.DecryptString(tempstr, myKey, True)
        Else
            DBUserName = ""
        End If
        tempstr = GetSetting("lab-i", "ORACLE", "Password", "")
        If Len(Trim(tempstr)) > 0 Then
            DBPassword = Blowfish.DecryptString(tempstr, myKey, True)
        Else
            DBPassword = ""
        End If
        tempstr = GetSetting("lab-i", "ORACLE", "Database", "")
        If Len(Trim(tempstr)) > 0 Then
            DBDatabase = Blowfish.DecryptString(tempstr, myKey, True)
        Else
            DBDatabase = ""
        End If
        tempstr = GetSetting("lab-i", "ORACLE", "Schema", "")
        If Len(Trim(tempstr)) > 0 Then
            DBSchema = Blowfish.DecryptString(tempstr, myKey, True)
        Else
            DBSchema = ""
        End If
        tempstr = GetSetting("lab-i", "ORACLE", "Server", Blowfish.EncryptString("(LOCAL)", myKey, True))
        If Len(Trim(tempstr)) > 0 Then
            DBServer = Blowfish.DecryptString(tempstr, myKey, True)
        Else
            DBServer = ""
        End If
        DBConnectionString = "Provider=MSDAORA;Data Source=" & DBServer & ";Data Source=" & DBDatabase & ";User Id=" & DBUserName & ";Password=" & DBPassword
End Select
End Sub

Public Sub OpenConnection()
On Error GoTo errFound
    myConnection.ConnectionString = DBConnectionString
    myConnection.ConnectionTimeout = 30
    myConnection.CursorLocation = adUseClient
    myConnection.Open
'    myConnection.Properties("Jet OLEDB:Max Locks Per File") = 500000
errFound:
    If Err.Number <> 0 Then
        Select Case Err.Number
            Case -2147467259
                MsgBox "Unable to connect to server" & vbCrLf & vbCrLf & "Please check database settings.", vbExclamation + vbOKOnly, App.ProductName
            Case Else
                MsgBox Err.Description & vbCrLf & "Please Contact your Systems Administrator.", vbExclamation + vbOKOnly, App.ProductName
        End Select
        isDBconnected = False
    Else
        isDBconnected = True
    End If
End Sub

Public Sub CloseConnection()
If isDBconnected = True Then
    myConnection.Close
    isDBconnected = False
End If
End Sub

Public Sub Highlight()
    With Screen.ActiveForm
        If (TypeOf .ActiveControl Is TextBox) Then
            .ActiveControl.SelStart = 0
            .ActiveControl.SelLength = Len(.ActiveControl)
        End If
    End With
End Sub

Public Function get_OtherResultInfo(mess As String, Optional asOBI As Integer = 0) As Variant
Dim SampleNum As String
Dim tmpVar As String
Dim tmpName As String
Dim tmpID As String
Dim tmpSequence As String
Dim ID1 As String
Dim ID2 As String
Dim tmpVar2(1 To 4) As String
Dim varSuspect(1 To 11) As String
Dim ctr As Integer
Dim startpoint As Integer
Dim strTemp As String
Dim cassPoss As String
    SampleNum = Trim(Str(Val(Mid(mess, 22, 12))))
    If Len(SampleNum) < 3 Then SampleNum = String(3 - Len(SampleNum), "0") & SampleNum
    strTemp = "ORI|" & InstrumentID & "|" 'Instrument ID
    ctr = 1
    tmpVar = ""
    cassPoss = ""
    If Len(Trim(Format(getMessageValues(mess, "PID "), ">"))) > 0 Or Len(Trim(Format(getMessageValues(mess, "PLN"), ">"))) > 0 Or Len(Trim(Format(getMessageValues(mess, "PFN"), ">"))) > 0 Then
        tmpID = Format(getMessageValues(mess, "PID "), ">")
        If InStr(1, tmpID, "-") > 0 Then
            tmpSequence = Mid(tmpID, InStr(1, tmpID, "-") + 1)
            tmpID = Mid(tmpID, 1, InStr(1, tmpID, "-") - 1)
            ID1 = tmpSequence
        Else
            ID1 = Format(getMessageValues(mess, "PID "), ">")
        End If
        ID2 = Format(getMessageValues(mess, "PLN"), ">") & ", " & Format(getMessageValues(mess, "PFN"), ">")
    Else
        ID1 = Format(getMessageValues(mess, "ID1"), ">")
        ID2 = Format(getMessageValues(mess, "ID2"), ">")
    End If
    cassPoss = Replace(getSTKsData(mess, "CASSPOS", "S", 1), "P", "")
    cassPoss = Replace(cassPoss, Chr(13), "")
    If cassPoss = "S" And InStr(ID1, "--") = 0 And Len(Trim(ID1)) = 0 Then
        tmpVar = ID1
    ElseIf cassPoss = "S" And InStr(ID1, "--") > 0 Then
        tmpVar = ID2
    ElseIf cassPoss = "S" And InStr(ID1, "--") = 0 And Len(Trim(ID1)) > 0 And Len(Trim(ID2)) > 0 Then
        tmpVar = ID2
    ElseIf cassPoss = "S" And InStr(ID1, "--") = 0 And Len(Trim(ID1)) > 0 And Len(Trim(ID2)) = 0 Then
        tmpVar = ID1
    ElseIf cassPoss <> "S" And Len(Trim(ID2)) < 1 And InStr(ID1, "--") > 0 Then
        tmpVar = cassPoss
    ElseIf cassPoss <> "S" And Len(Trim(ID2)) > 0 Then
        tmpVar = ID2
    End If
    
    If cassPoss <> "S" And Len(Trim(Format(getMessageValues(mess, "PID"), ">"))) > 0 Then
        tmpVar = Format(getMessageValues(mess, "ID1"), ">")
    End If
'    ID1 = getSTKsData(mess, "ID1", getSTKsData(mess, "NAME", ""))
'    ID2 = getSTKsData(mess, "ID2", getSTKsData(mess, "SEQUENCE", ""))
'    If Len(ID2) > 0 Then tmpVar = ID2 Else tmpVar = ID1
    If Len(Trim(ID2)) > 0 Then
        strTemp = strTemp & Trim(ID2) & "|" 'Patient Name / ID
    Else
        strTemp = strTemp & Trim(tmpVar) & "|" 'Patient Name / ID
    End If
    'tmpName = Trim(tmpVar)
    If IsNumeric(ID1) Then
        tmpVar = ID1
    ElseIf IsNumeric(ID2) Then
        tmpVar = ID2
    Else
        tmpVar = ""
    End If
'    If getMessageValues(mess, "CASSPOS") = "S" Then tmpVar = ID2
    If Len(Trim(ID2)) > 0 Then
        strTemp = strTemp & Trim(getSTKsData(mess, "ID2", getSTKsData(mess, "SEQUENCE", ""))) & "|" 'Sample ID Number
    Else
        strTemp = strTemp & Trim(tmpVar) & "|" 'Sample ID Number
    End If
    strTemp = strTemp & "" & "|" 'Rack Number
    strTemp = strTemp & "" & "|" 'Tube Position
    If InStr(1, mess, "SEQUENCE") > 0 Then
        strTemp = strTemp & getMessageValues(mess, "SEQUENCE") & "|" 'Sequence Number
    Else
        strTemp = strTemp & tmpVar & "|" 'Sequence Number
    End If
    strTemp = strTemp & getMessageValues(mess, "PID ") & "|" 'ID Information
    strTemp = strTemp & getMessageValues(mess, "BIRTH") & "|"  'Reserve 1 - Birthdate
    strTemp = strTemp & "" & "|" 'Reserve 2 - Room
    strTemp = strTemp & "" & "|" 'Reserve 3 - Age
    strTemp = strTemp & getMessageValues(mess, "SEX") & "|"  'Reserve 4 - Sex
    strTemp = strTemp & "" & "|" 'Reserve 5 - Specimen
    strTemp = strTemp & getDefinitive(mess) & "|" 'Reserve 6 - definitive
    strTemp = strTemp & "" & "|" 'Reserve 7 - Doctor
    strTemp = strTemp & "" & "|" 'Reserve 8
    If InStr(1, mess, "CONDITION") > 0 Then
        startpoint = InStr(1, mess, "CONDITION")
        tmpVar2(1) = 0: tmpVar2(2) = 0: tmpVar2(3) = 0: tmpVar2(4) = 0
        Do While startpoint > 0
            tmpVar = ""
            ctr = 1
            Do While Mid(mess, InStr(startpoint, mess, "CONDITION") + 9 + ctr, 1) <> Chr(13) And ctr <= Len(mess)
                tmpVar = tmpVar + Mid(mess, InStr(startpoint, mess, "CONDITION") + 9 + ctr, 1)
                ctr = ctr + 1
            Loop
            Select Case Trim(Format(tmpVar, ">"))
                Case "ABNORMAL WBC POP": tmpVar2(1) = "0"
                Case "ABNORMAL RBC POP": tmpVar2(2) = "0"
                Case "ABNORMAL PLT POP": tmpVar2(3) = "0"
                Case "VERIFY RETIC": tmpVar2(4) = "0"
            End Select
            startpoint = InStr(startpoint + 1, mess, "CONDITION")
        Loop
    End If
    tmpVar = tmpVar2(1) & tmpVar2(2) & tmpVar2(3) & tmpVar2(4)
    strTemp = strTemp & tmpVar & "|" 'Reserve 9
    strTemp = strTemp & "|" 'Reserve 10

    If InStr(1, mess, "SUSPECT") > 0 Then
        startpoint = InStr(1, mess, "SUSPECT")
        For ctr = 1 To 11
        varSuspect(ctr) = 0
    Next ctr
        Do While startpoint > 0
            tmpVar = ""
            ctr = 1
            Do While Mid(mess, InStr(startpoint, mess, "SUSPECT") + 7 + ctr, 1) <> Chr(13)
                tmpVar = tmpVar + Mid(mess, InStr(startpoint, mess, "SUSPECT") + 7 + ctr, 1)
                ctr = ctr + 1
            Loop
            Select Case Trim(Format(tmpVar, ">"))
                Case "BLASTS": varSuspect(1) = "1"
                Case "IMM GRANS/BANDS (1 OR 2)": varSuspect(2) = "1"
                Case "VARIANT LYMPHS": varSuspect(3) = "1"
                Case "REVIEW SLIDE": varSuspect(4) = "1"
                Case "NRBCS": varSuspect(5) = "1"
                Case "DIMORPHIC RBC POP": varSuspect(6) = "1"
                Case "MICRO RBCS/": varSuspect(7) = "1"
                Case "RBC FRAGMENTS": varSuspect(8) = "1"
                Case "RBC AGGLUTINATION": varSuspect(9) = "1"
                Case "PLATELET CLUMPS": varSuspect(10) = "1"
                Case "GIANT PLATELETS": varSuspect(11) = "1"
            End Select
            startpoint = InStr(startpoint + 1, mess, "SUSPECT")
        Loop
        tmpVar = ""
        For ctr = 1 To 11
            tmpVar = tmpVar & varSuspect(ctr)
        Next ctr
    End If
    tmpVar = ""
    If asOBI = 0 Then
        strTemp = strTemp & tmpVar 'Reserve 11
        get_OtherResultInfo = Split(strTemp, "|")
    ElseIf asOBI = 1 Then
        strTemp = strTemp & tmpVar 'Reserve 11
        get_OtherResultInfo = strTemp
    Else
        strTemp = strTemp & tmpVar 'Reserve 11
        get_OtherResultInfo = Split(strTemp, "|")
    End If
End Function

Public Function segment_OBX2(mess As String, rstExams As ADODB.Recordset) As String
Dim strTemp As String
Dim strDate As String
Dim tempResult As String
Dim ctr As Integer
Dim strEndString As String
Dim rstFormat As ADODB.Recordset
    If InStr(1, mess, "DATE") > 0 Then
        strDate = Mid(mess, InStr(1, mess, "DATE") + 5, 8) & " " & Mid(mess, InStr(1, mess, "TIME") + 5, 8)
        If Not IsDate(strDate) Then strDate = CStr(Now)
    Else
        strDate = CStr(Now)
    End If
    strDate = Format(strDate, "YYYYMMddHHmmss")
    strEndString = "||||||" & strDate & "|" & InstrumentID & vbCr
    strTemp = ""
    If rstExams.State = 1 Then
        If rstExams.RecordCount > 0 Then
            rstExams.MoveFirst
            For ctr = 0 To rstExams.RecordCount - 1
                Set rstFormat = myConnection.Execute("Select * from tbLabProfileExams where ProfileID = '" & DefaultProfileID & "' and InstrumentID = '" & InstrumentID & "' and ExamID = '" & rstExams.Fields("ExamID") & "'")
                If rstFormat.RecordCount > 0 Then
                    tempResult = getStrResult(1, mess, rstExams.Fields("InstExamName").Value)
                    If IsNumeric(tempResult) Then
                        strTemp = strTemp & "OBX|" & ctr & "|ST|" & rstExams.Fields("InstExamName").Value & "||" & Format(CDbl(tempResult) * rstFormat.Fields("Dilution").Value, rstFormat.Fields("DataFormat").Value) & "|" & rstExams.Fields("Unit").Value & "||" & getStrResult(2, mess, rstExams.Fields("InstExamName").Value) & strEndString
                    End If
                Else
                    tempResult = getStrResult(1, mess, rstExams.Fields("InstExamName").Value)
                    If IsNumeric(tempResult) Then
                        strTemp = strTemp & "OBX|" & ctr & "|ST|" & rstExams.Fields("InstExamName").Value & "||" & getStrResult(1, mess, rstExams.Fields("InstExamName").Value) & "|" & rstExams.Fields("Unit").Value & "||" & getStrResult(2, mess, rstExams.Fields("InstExamName").Value) & strEndString
                    End If
                End If
                If Not rstExams.EOF Then rstExams.MoveNext
            Next ctr
        End If
    End If
    Set rstFormat = Nothing
    segment_OBX2 = strTemp
End Function

Public Function segment_OBX(mess As String, rstExams As ADODB.Recordset) As String
Dim strTemp As String
Dim strDate As String
Dim ctr As Integer
Dim strEndString As String
    If InStr(1, mess, "DATE") > 0 Then
        strDate = Mid(mess, InStr(1, mess, "DATE") + 5, 8) & " " & Mid(mess, InStr(1, mess, "TIME") + 5, 8)
        If Not IsDate(strDate) Then strDate = CStr(Now)
    Else
        strDate = CStr(Now)
    End If
    If Year(strDate) = "1899" Then strDate = Format(Now, "mm/dd/yyyy") & " " & Format(strDate, "HH:mm:ss")
    strDate = Format(strDate, "YYYYMMddHHmmss")
    strEndString = "||||||" & strDate & "|" & InstrumentID & vbCr
    strTemp = ""
    If rstExams.State = 1 Then
        If rstExams.RecordCount > 0 Then
            rstExams.MoveFirst
            For ctr = 0 To rstExams.RecordCount - 1
                strTemp = strTemp & "OBX|" & ctr & "|ST|" & rstExams.Fields("InstExamName").Value & "||" & getStrResult(1, mess, rstExams.Fields("InstExamName").Value) & "|" & rstExams.Fields("Unit").Value & "||" & getStrResult(2, mess, rstExams.Fields("InstExamName").Value) & strEndString
                If Not rstExams.EOF Then rstExams.MoveNext
            Next ctr
        End If
    End If
    segment_OBX = strTemp
End Function

Public Function segment_OBX_UC(mess As String, rstExams As ADODB.Recordset) As String
Dim strTemp As String
Dim strDate As String
Dim ctr As Integer
Dim strEndString As String
    If InStr(1, mess, "DATE") > 0 Then
        strDate = Mid(mess, InStr(1, mess, "DATE") + 5, 8) & " " & Mid(mess, InStr(1, mess, "TIME") + 5, 8)
    Else
        strDate = CStr(Now)
    End If
    strDate = Format(strDate, "YYYYMMddHHmmss")
    strEndString = "||||||" & strDate & "|" & InstrumentID & vbCr
    strTemp = ""
    If rstExams.State = 1 Then
        If rstExams.RecordCount > 0 Then
            rstExams.MoveFirst
            For ctr = 0 To rstExams.RecordCount - 1
                strTemp = strTemp & "OBX|" & ctr & "|ST|" & rstExams.Fields("InstExamName").Value & "||" & getStrResult(1, mess, rstExams.Fields("InstExamName").Value) & "|" & rstExams.Fields("Unit").Value & "||" & getStrResult(2, mess, rstExams.Fields("InstExamName").Value) & strEndString
                If Not rstExams.EOF Then rstExams.MoveNext
            Next ctr
        End If
    End If
    segment_OBX_UC = strTemp
End Function


Private Function getStrResult(mArg As Byte, mess As String, searchString As String)
Dim myValues(1 To 2) As String
Dim partString As String
Dim startpoint As Integer
Dim ctr As Integer
myValues(1) = ""
myValues(2) = ""
partString = ""
ctr = 0
startpoint = InStr(1, mess, searchString)
    If startpoint > 0 Then
        Do While (Mid(mess, startpoint + ctr, 1) <> Chr(13) And startpoint + ctr <= Len(mess))
            partString = partString + Mid(mess, startpoint + ctr, 1)
            ctr = ctr + 1
        Loop
        partString = Trim(Replace(partString, searchString, ""))
        ctr = 1
        Do While Mid(partString, ctr, 1) <> " " And ctr <= Len(partString)
            myValues(1) = myValues(1) & Mid(partString, ctr, 1)
            ctr = ctr + 1
        Loop
        partString = Trim(Replace(partString, myValues(1), ""))
        ctr = 1
        If Mid(partString, ctr, 1) <> " " And ctr <= Len(partString) Then
            Do While Mid(partString, ctr + 1, 1) <> " " And ctr <= Len(partString)
                myValues(2) = myValues(2) & Mid(partString, ctr, 1)
                ctr = ctr + 1
            Loop
        End If
    End If
    getStrResult = myValues(mArg)
End Function

Public Function getMessageValues(mess As String, searchString As String, Optional asInt As Boolean = False) As String
Dim partString As String
Dim startpoint As Integer
Dim ctr As Integer
partString = ""
ctr = 0
startpoint = InStr(1, mess, searchString)
    If startpoint > 0 Then
        Do While Mid(mess, startpoint + ctr, 1) <> Chr(4)
            partString = partString + Mid(mess, startpoint + ctr, 1)
            ctr = ctr + 1
        Loop
        partString = Replace(partString, searchString, "")
        partString = Trim$(Replace(partString, Chr(13), ""))
    End If
    getMessageValues = partString
    If asInt Then
        If IsNumeric(getMessageValues) Then getMessageValues = Val(getMessageValues)
    End If
End Function

Private Function getDefinitive(myMess As String) As String
Dim mFlags(1 To 38) As String
Dim tempstr As String
Dim ctr As Integer
Dim startpoint As Integer
For ctr = 1 To 38
    mFlags(ctr) = 0
Next ctr
    If InStr(1, myMess, "DEFINIT") > 0 Then
        startpoint = InStr(1, myMess, "DEFINIT")
        Do While startpoint > 0
            tempstr = ""
            ctr = 1
            Do While Mid(myMess, InStr(startpoint, myMess, "DEFINIT") + 7 + ctr, 1) <> Chr(13)
                tempstr = tempstr + Mid(myMess, InStr(startpoint, myMess, "DEFINIT") + 7 + ctr, 1)
                ctr = ctr + 1
            Loop
            Select Case Trim(Format(tempstr, ">"))
                Case "LEUKOPENIA": mFlags(1) = "1"
                Case "LEUKOCYTOSIS": mFlags(2) = "1"
                Case "NEUTROPENIA %": mFlags(3) = "1"
                Case "NEUTROPENIA #": mFlags(4) = "1"
                Case "NEUTROPHILIA %": mFlags(5) = "1"
                Case "NEUTROPHILIA #": mFlags(6) = "1"
                Case "LYMPHOPENIA %": mFlags(7) = "1"
                Case "LYMPHOPENIA #": mFlags(8) = "1"
                Case "LYMPHOCYTOSIS %": mFlags(9) = "1"
                Case "LYMPHOCYTOSIS #": mFlags(10) = "1"
                Case "MONOCYTOSIS %": mFlags(11) = "1"
                Case "MONOCYTOSIS #": mFlags(12) = "1"
                Case "EOSINOPHILIA %": mFlags(13) = "1"
                Case "EOSINOPHILIA #": mFlags(14) = "1"
                Case "BASOPHILIA %": mFlags(15) = "1"
                Case "BASOPHILIA #": mFlags(16) = "1"
                Case "ANEMIA": mFlags(17) = "1"
                Case "1+ ANISOCYTOSIS": mFlags(18) = "1"
                Case "2+ ANISOCYTOSIS ": mFlags(19) = "1"
                Case "3+ ANISOCYTOSIS": mFlags(20) = "1"
                Case "1+ MICROCYTOSIS": mFlags(21) = "1"
                Case "2+ MICROCYTOSIS": mFlags(22) = "1"
                Case "3+ MICROCYTOSIS": mFlags(23) = "1"
                Case "1+ MACROCYTOSIS": mFlags(24) = "1"
                Case "2+ MACROCYTOSIS": mFlags(25) = "1"
                Case "3+ MACROCYTOSIS": mFlags(26) = "1"
                Case "1+ HYPOCHROMIA": mFlags(27) = "1"
                Case "2+ HYPOCHROMIA": mFlags(28) = "1"
                Case "3+ HYPOCHROMIA": mFlags(29) = "1"
                Case "1+ POIKILOCYTOSIS": mFlags(30) = "1"
                Case "2+ POIKILOCYTOSIS": mFlags(31) = "1"
                Case "3+ POIKILOCYTOSIS": mFlags(32) = "1"
                Case "ERYTHROCYTOSIS": mFlags(33) = "1"
                Case "PANCYTOPENIA": mFlags(34) = "1"
                Case "THROMBOCYTOPENIA": mFlags(35) = "1"
                Case "THROMBOCYTOSIS": mFlags(36) = "1"
                Case "SMALL PLATELETS": mFlags(37) = "1"
                Case "LARGE PLATELETS": mFlags(38) = "1"
            End Select
            startpoint = InStr(startpoint + 1, myMess, "DEFINIT")
        Loop
        tempstr = ""
    End If
    For ctr = 1 To 38
        tempstr = tempstr & mFlags(ctr)
    Next ctr
    getDefinitive = tempstr
End Function

Public Function formatDT(DTstr As Variant, myDTdef As Integer) As String
Dim Temp(1 To 6) As String
Dim tempDate As String
Dim tempTime As String
Dim tempstr As String
Temp(1) = Mid(DTstr, 13, 2)
Temp(2) = Mid(DTstr, 11, 2)
Temp(3) = Mid(DTstr, 9, 2)
Temp(4) = Mid(DTstr, 7, 2)
Temp(5) = Mid(DTstr, 5, 2)
Temp(6) = Mid(DTstr, 1, 4)
tempDate = Temp(5) & "/" & Temp(4) & "/" & Temp(6)
tempTime = Temp(3) & ":" & Temp(2) & ":" & Temp(1)
If Not IsDate(tempDate & " " & tempTime) Then
    tempDate = Format(Now, "mm/dd/yyyy")
    tempTime = Format(Now, "h:mm:ss AMPM")
End If
Select Case myDTdef
    Case 1
        formatDT = Format$(tempDate, "mmmm dd, yyyy")
    Case 2
        formatDT = Format(tempTime, "h:mm:ss AMPM")
    Case 3
        formatDT = Format(tempDate, "mmmm dd, yyyy") & " " & Format(tempTime, "h:mm:ss AMPM")
End Select
End Function

Public Function getStringData(mess As String, searchTag As String, DefaultValue As String, Optional colCount As Byte = 1) As String
Dim tempstr As String
Dim tempColCount As Integer
Dim StartCount As Long
Dim foundData As Boolean
Dim dataEnd As Boolean
foundData = False
dataEnd = False
tempstr = ""
tempColCount = 0
If IsNumeric(colCount) = False Then colCount = 1
StartCount = InStr(1, mess, searchTag)
If StartCount > 0 Then
    StartCount = StartCount + Len(searchTag)
    Do While tempColCount <> colCount And StartCount <= Len(mess)
        Do While (foundData = False Or dataEnd = False)
            If dataEnd = False Then tempstr = tempstr & Mid(mess, StartCount, 1)
            If foundData = True Then
                If Mid(mess, StartCount, 1) = " " Or Mid(mess, StartCount, 1) = vbCrLf Or Mid(mess, StartCount, 1) = vbCr Or StartCount >= Len(mess) Then dataEnd = True
            Else
                If Len(Trim(Mid(mess, StartCount, 1))) > 0 Then
                    foundData = True
                Else
                    If StartCount = Len(mess) Then
                        foundData = True
                        dataEnd = True
                    End If
                End If
            End If
            StartCount = StartCount + 1
        Loop
        tempColCount = tempColCount + 1
        If tempColCount <> colCount Then
            tempstr = ""
            foundData = False
            dataEnd = False
        Else
            tempstr = Trim(tempstr)
        End If
    Loop
Else
    tempstr = DefaultValue
End If
If Len(Trim(tempstr)) = 0 Then tempstr = DefaultValue
getStringData = tempstr
End Function

Public Function DisplayResult2Text(msgHL7format As String, OtherInfo As Variant) As String
Dim ctr1 As Integer
Dim counter As Integer
Dim myResults As Variant
Dim myResults2 As Variant
Dim tempstr As String
Dim strSql As String
On Error GoTo errFound
myResults = Split(msgHL7format, vbCr)
counter = UBound(myResults) - 1
DisplayResult2Text = DisplayResult2Text & "Exam          Result        Unit         " & vbCrLf
DisplayResult2Text = DisplayResult2Text & "------------- ------------- -------------" & vbCrLf
For ctr1 = 0 To counter
    myResults2 = Split(myResults(ctr1), "|")
    DisplayResult2Text = DisplayResult2Text _
            & myResults2(3) & String(14 - Len(myResults2(3)), " ") _
            & myResults2(5) & String(14 - Len(myResults2(5)), " ") _
            & myResults2(6) & String(14 - Len(myResults2(6)), " ") & vbCrLf
            '& myResults2(8) & vbCrLf
            '& myMsgDef.get_Abnormalflags(Val(myResults2(8))) & vbCrLf
Next ctr1
tempstr = "Sample Number: " & OtherInfo(3) & vbCrLf & _
            "Resultdate:" & formatDT(myResults2(14), 1) & vbCrLf
DisplayResult2Text = tempstr & DisplayResult2Text & String(41, "-") & vbCrLf
errFound:
    If Err.Number <> 0 Then DisplayResult2Text = "Conversion Error..."
End Function

Public Function getSTKsData(mess As String, searchTag As String, DefaultValue As String, Optional colCount As Byte = 1) As String
Dim tempstr As String
Dim tempColCount As Integer
Dim StartCount As Long
Dim foundData As Boolean
Dim dataEnd As Boolean
foundData = False
dataEnd = False
tempstr = ""
tempColCount = 0
If IsNumeric(colCount) = False Then colCount = 1
StartCount = InStr(1, mess, searchTag)
If StartCount > 0 Then
    StartCount = StartCount + Len(searchTag)
    Do While tempColCount <> colCount And StartCount <= Len(mess)
        Do While (foundData = False Or dataEnd = False)
            If dataEnd = False Then tempstr = tempstr & Mid(mess, StartCount, 1)
            If foundData = True Then
                If Mid(mess, StartCount, 1) = " " Or Mid(mess, StartCount, 1) = vbCrLf Or Mid(mess, StartCount, 1) = vbCr Or StartCount >= Len(mess) Then dataEnd = True
            Else
                If Len(Trim(Mid(mess, StartCount, 1))) > 0 Then
                    foundData = True
                Else
                    If StartCount = Len(mess) Then
                        foundData = True
                        dataEnd = True
                    End If
                End If
            End If
            StartCount = StartCount + 1
        Loop
        tempColCount = tempColCount + 1
        If tempColCount <> colCount Then
            tempstr = ""
            foundData = False
            dataEnd = False
        Else
            tempstr = Trim(tempstr)
        End If
    Loop
Else
    tempstr = DefaultValue
End If
If Len(Trim(tempstr)) = 0 Then tempstr = DefaultValue
getSTKsData = tempstr
End Function

Public Function ScalePictureBoxImage(oSrcPic As MDIForm) As Boolean
Dim lDesktopHwnd As Long
Dim lDesktopDC As Long
Dim lTempDC As Long
Dim lPrevDIB As Long
Dim bIsAutoDraw As Boolean
Dim lPrevStretchBltMode As Long
Dim SrcX As Long
Dim SrcY As Long
Dim DestX As Long
Dim DestY As Long

If (oSrcPic.Picture = 0) Then Exit Function 'No picture to stretch!
'Get a handle to the desktop's DC and create a temporary DC
lDesktopHwnd = GetDesktopWindow()
lDesktopDC = GetDC(lDesktopHwnd)
lTempDC = CreateCompatibleDC(lDesktopDC)
ReleaseDC lDesktopHwnd, lDesktopDC
    If lTempDC Then
    ' Copy the PictureBox's Picture to the DC we created
        lPrevDIB = SelectObject(lTempDC, oSrcPic.Picture.handle)
        If lPrevDIB Then
            With oSrcPic
            ' Calculate the source and destination sizes
                SrcX = .ScaleX(.Picture.Width, vbHimetric, vbPixels)
                SrcY = .ScaleY(.Picture.Height, vbHimetric, vbPixels)
                DestX = .ScaleX(.ScaleWidth, .ScaleMode, vbPixels)
                DestY = .ScaleY(.ScaleHeight, .ScaleMode, vbPixels)
                ' Store the current AutoRedraw setting and change it to true
                bIsAutoDraw = .AutoRedraw
                .AutoRedraw = True
                ' This helps avoid destruction of the image
                lPrevStretchBltMode = SetStretchBltMode(.hdc, COLORONCOLOR)
                ' Copy the image from the temporary DC to the PictureBox using the
                ' new size
                StretchBlt .hdc, 0, 0, DestX, DestY, lTempDC, 0, 0, SrcX, SrcY, vbSrcCopy
                ' Reset the StretchBlt mode
                SetStretchBltMode .hdc, lPrevStretchBltMode
                ' Dump the DC we created by replacing it with the original
                SelectObject lTempDC, lPrevDIB
                ' Refresh the picture box
                .Refresh
                ' Reset the AutoRedraw
                .AutoRedraw = bIsAutoDraw
            End With
            ' Clean up
            DeleteDC lTempDC
            ' Pass back a value indicating everything is ok
            ScalePictureBoxImage = True
            ' Jump out so we don't pass back the False value at
            ' the end of the function
            Exit Function
        Else
        ' We got a DC but no DIB so clean it up
        DeleteDC lTempDC
        End If
    End If
    ' If either of the If..Then evaluate to False, we'll end up here
    ScalePictureBoxImage = False
End Function

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

Public Sub setButtons(mConnect As Boolean, mDisconnect As Boolean, _
    mLabexam As Boolean, mCommunication As Boolean, mConnection As Boolean)
With mdiMain
    .mnuConnect.Enabled = mConnect
    .mnuSendReceive.Enabled = mConnect
    .mnuSendID.Enabled = mConnect
    .tbrMain.Buttons("Connect").Enabled = mConnect
    .mnuDisconnect.Enabled = mDisconnect
    .tbrMain.Buttons("Disconnect").Enabled = mDisconnect
    .mnuComSet.Enabled = mCommunication
    .tbrMain.Buttons("Communication").Enabled = mCommunication
    .mnuConSet.Enabled = mConnection
    .mnuLabExam.Enabled = mLabexam
    .tbrMain.Buttons("Lab Exams").Enabled = mLabexam
End With
End Sub

Public Sub LinkTable(psTable As String, psFromPath As String, _
  psToPath As String, Optional newName As String = "", Optional mPassword As String = "")
Dim cat As ADOX.Catalog
Dim tbl As ADOX.Table
Dim sShortPath As String
Dim ctr As Integer
Dim tbFound As Boolean
On Error GoTo errFound
    'get short path name of the source database
    sShortPath = Space(255)
    Call GetShortPathName(psFromPath, sShortPath, 255)
    sShortPath = Trim$(sShortPath)
    sShortPath = Left$(sShortPath, Len(sShortPath) - 1)
    tbFound = False
    
    Set cat = New ADOX.Catalog
    Set cat.ActiveConnection = myConnection
    
    'link table
    Set tbl = New ADOX.Table
    With tbl
         If Len(Trim(newName)) > 0 Then
            .name = newName
         Else
            .name = psTable
         End If
         Set .ParentCatalog = cat
         .Properties("Jet OLEDB:Create Link") = True
         .Properties("Jet OLEDB:Link Datasource") = sShortPath
         .Properties("Jet OLEDB:Remote Table Name") = psTable
         If Len(Trim(mPassword)) > 0 Then
            .Properties("Jet OLEDB:Link Provider String") = ";Pwd=" & mPassword
         Else
            If Len(Trim(DBPassword)) > 0 Then .Properties("Jet OLEDB:Link Provider String") = ";Pwd=" & DBPassword
         End If
         For ctr = 0 To cat.Tables.Count - 1
            If Format(cat.Tables(ctr).name, ">") = Format(.name, ">") Then
                tbFound = True
                Exit For
            Else
                tbFound = False
            End If
         Next ctr
         If tbFound Then cat.Tables.Delete .name
         cat.Tables.Append tbl
    End With
    Set tbl = Nothing
    Set cat = Nothing
errFound:
    If Err.Number = -2147217860 Then
        MsgBox "Table: " & psTable & " not found. Press OK to continue", vbOKOnly, "Error"
        Resume Next
    ElseIf Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description, vbOKOnly + vbExclamation, "Error"
    End If
End Sub

Public Sub FixLink(mDataLocation As String, Optional Check As Boolean = True)
Dim TempLocation As String
TempLocation = Replace(Format(mDataLocation, ">"), "LABORATORY.MDB", "")
If Check Then
    If GetSetting("Lab-I", "DataLocation", "LastPath", "") <> TempLocation Then
        LinkTable "tbPatientMaster", TempLocation & "LabPatientData.mdb", mDataLocation
        LinkTable "tbPatientValues", TempLocation & "LabPatientValues.mdb", mDataLocation
        LinkTable "tbLabInstrumentResults", TempLocation & "InstrumentResults.mdb", mDataLocation
        LinkTable "tbLabInstrumentQC", TempLocation & "InstrumentResults.mdb", mDataLocation
        LinkTable "LabUnprocessed", TempLocation & "LabBuffer.mdb", mDataLocation
        LinkTable "LabTempPrint", TempLocation & "LabBuffer.mdb", mDataLocation
        LinkTable "tbLabInstrumentRValues", TempLocation & "InstrumentRValues.mdb", mDataLocation
        LinkTable "tbLabRequest", TempLocation & "LabRequestBuffer.mdb", mDataLocation
        LinkTable "tbLabRequestForm", TempLocation & "LabRequestForm.mdb", mDataLocation
        SaveSetting "LAB-I", "DataLocation", "LastPath", TempLocation
    End If
Else
    LinkTable "tbPatientMaster", TempLocation & "LabPatientData.mdb", mDataLocation
    LinkTable "tbPatientValues", TempLocation & "LabPatientValues.mdb", mDataLocation
    LinkTable "tbLabInstrumentQC", TempLocation & "InstrumentResults.mdb", mDataLocation
    LinkTable "LabUnprocessed", TempLocation & "LabBuffer.mdb", mDataLocation
    LinkTable "LabTempPrint", TempLocation & "LabBuffer.mdb", mDataLocation
    LinkTable "tbLabInstrumentRValues", TempLocation & "InstrumentRValues.mdb", mDataLocation
    LinkTable "tbLabInstrumentResults", TempLocation & "InstrumentResults.mdb", mDataLocation
    LinkTable "tbLabRequest", TempLocation & "LabRequestBuffer.mdb", mDataLocation
    LinkTable "tbLabRequestForm", TempLocation & "LabRequestForm.mdb", mDataLocation
    SaveSetting "Lab-I", "DataLocation", "LastPath", TempLocation
End If
End Sub

Public Sub setResultReport()
    RPTName = GetSetting("HostForms_SA", "Settings", "ReportForm", "")
End Sub

Public Sub setHostForms()
    Set HostSA = CreateObject("HostForms_SA.ClassForms")
    HostSA.wDBServerType = DBServerType
    HostSA.wDBUserName = DBUserName
    HostSA.wDBPassword = DBPassword
    HostSA.wDBDatabase = DBDatabase
    HostSA.wDBSchema = DBSchema
    HostSA.wDBServer = DBServer
    HostSA.wDBConnectionString = DBConnectionString
    HostSA.wConnection = myConnection
    HostSA.wInstrumentID = InstrumentID
    HostSA.wRPTname = RPTName
    HostSA.wSectionCode = SectionCode
End Sub

Public Sub BackupDatabase(sFolder As String, Optional confirmOverwrite As Boolean = True)
Dim strDir As String
Dim strCurDir As String
Dim strSrcDir As String
DoEvents
CloseConnection
strCurDir = Replace(Format(DBDatabase, ">"), "LABORATORY.MDB", "")
strSrcDir = strCurDir
strCurDir = strCurDir & "Backup"
strDir = Dir(strCurDir, vbDirectory)
If Len(Trim(strDir)) = 0 Then
    fsoFiles.Createfolder strCurDir
End If
strDir = Dir(strCurDir & "\" & sFolder, vbDirectory)
If Len(Trim(strDir)) = 0 Then
    fsoFiles.Createfolder strCurDir & "\" & sFolder
End If
strCurDir = strCurDir & "\" & sFolder & "\"
'**************************************
If Len(Trim(Dir(strCurDir & "Laboratory.mdb", vbNormal))) > 0 Then
    If confirmOverwrite Then
        If MsgBox("Do you want to overwrite the existing Backup?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
            fsoFiles.CopyFile strSrcDir & "Laboratory.mdb", strCurDir & "Laboratory.mdb", True
            fsoFiles.CopyFile strSrcDir & "InstrumentResults.mdb", strCurDir & "InstrumentResults.mdb", True
            fsoFiles.CopyFile strSrcDir & "LabBuffer.mdb", strCurDir & "LabBuffer.mdb", True
            fsoFiles.CopyFile strSrcDir & "LabPatientData.mdb", strCurDir & "LabPatientData.mdb", True
            fsoFiles.CopyFile strSrcDir & "LabPatientValues.mdb", strCurDir & "LabPatientValues.mdb", True
            fsoFiles.CopyFile strSrcDir & "InstrumentRValues.mdb", strCurDir & "InstrumentRValues.mdb", True
            DAO.DBEngine.CompactDatabase strSrcDir & "Laboratory.mdb", strSrcDir & "tLaboratory.mdb", dbLangGeneral, , ";pwd=" & DBPassword
            DAO.DBEngine.CompactDatabase strSrcDir & "InstrumentResults.mdb", strSrcDir & "tInstrumentResults.mdb", dbLangGeneral, , ";pwd=" & DBPassword
            DAO.DBEngine.CompactDatabase strSrcDir & "LabBuffer.mdb", strSrcDir & "tLabBuffer.mdb", dbLangGeneral, , ";pwd=" & DBPassword
            DAO.DBEngine.CompactDatabase strSrcDir & "LabPatientData.mdb", strSrcDir & "tLabPatientData.mdb", dbLangGeneral, , ";pwd=" & DBPassword
            DAO.DBEngine.CompactDatabase strSrcDir & "LabPatientValues.mdb", strSrcDir & "tLabPatientValues.mdb", dbLangGeneral, , ";pwd=" & DBPassword
            DAO.DBEngine.CompactDatabase strSrcDir & "InstrumentRValues.mdb", strSrcDir & "tInstrumentRValues.mdb", dbLangGeneral, , ";pwd=" & DBPassword
            fsoFiles.CopyFile strSrcDir & "tLaboratory.mdb", strSrcDir & "Laboratory.mdb", True
            fsoFiles.CopyFile strSrcDir & "tInstrumentResults.mdb", strSrcDir & "InstrumentResults.mdb", True
            fsoFiles.CopyFile strSrcDir & "tLabBuffer.mdb", strSrcDir & "LabBuffer.mdb", True
            fsoFiles.CopyFile strSrcDir & "tLabPatientData.mdb", strSrcDir & "LabPatientData.mdb", True
            fsoFiles.CopyFile strSrcDir & "tLabPatientValues.mdb", strSrcDir & "LabPatientValues.mdb", True
            fsoFiles.CopyFile strSrcDir & "tInstrumentRValues.mdb", strSrcDir & "InstrumentRValues.mdb", True
            fsoFiles.DeleteFile strSrcDir & "tLaboratory.mdb", True
            fsoFiles.DeleteFile strSrcDir & "tInstrumentResults.mdb", True
            fsoFiles.DeleteFile strSrcDir & "tLabBuffer.mdb", True
            fsoFiles.DeleteFile strSrcDir & "tLabPatientData.mdb", True
            fsoFiles.DeleteFile strSrcDir & "tLabPatientValues.mdb", True
            fsoFiles.DeleteFile strSrcDir & "tInstrumentRValues.mdb", True
            MsgBox "Backup Successful. Press Ok to continue.", vbInformation + vbOKOnly, App.ProductName
        Else
            MsgBox "Unable to Backup Database. Press Ok to continue.", vbExclamation + vbOKOnly, App.ProductName
        End If
    Else
        fsoFiles.CopyFile strSrcDir & "Laboratory.mdb", strCurDir & "Laboratory.mdb", True
        fsoFiles.CopyFile strSrcDir & "InstrumentResults.mdb", strCurDir & "InstrumentResults.mdb", True
        fsoFiles.CopyFile strSrcDir & "LabBuffer.mdb", strCurDir & "LabBuffer.mdb", True
        fsoFiles.CopyFile strSrcDir & "LabPatientData.mdb", strCurDir & "LabPatientData.mdb", True
        fsoFiles.CopyFile strSrcDir & "LabPatientValues.mdb", strCurDir & "LabPatientValues.mdb", True
        fsoFiles.CopyFile strSrcDir & "InstrumentRValues.mdb", strCurDir & "InstrumentRValues.mdb", True
    End If
Else
    fsoFiles.CopyFile strSrcDir & "Laboratory.mdb", strCurDir & "Laboratory.mdb", True
    fsoFiles.CopyFile strSrcDir & "InstrumentResults.mdb", strCurDir & "InstrumentResults.mdb", True
    fsoFiles.CopyFile strSrcDir & "LabBuffer.mdb", strCurDir & "LabBuffer.mdb", True
    fsoFiles.CopyFile strSrcDir & "LabPatientData.mdb", strCurDir & "LabPatientData.mdb", True
    fsoFiles.CopyFile strSrcDir & "LabPatientValues.mdb", strCurDir & "LabPatientValues.mdb", True
    fsoFiles.CopyFile strSrcDir & "InstrumentRValues.mdb", strCurDir & "InstrumentRValues.mdb", True
    DAO.DBEngine.CompactDatabase strSrcDir & "Laboratory.mdb", strSrcDir & "tLaboratory.mdb", dbLangGeneral, , ";pwd=" & DBPassword
    DAO.DBEngine.CompactDatabase strSrcDir & "InstrumentResults.mdb", strSrcDir & "tInstrumentResults.mdb", dbLangGeneral, , ";pwd=" & DBPassword
    DAO.DBEngine.CompactDatabase strSrcDir & "LabBuffer.mdb", strSrcDir & "tLabBuffer.mdb", dbLangGeneral, , ";pwd=" & DBPassword
    DAO.DBEngine.CompactDatabase strSrcDir & "LabPatientData.mdb", strSrcDir & "tLabPatientData.mdb", dbLangGeneral, , ";pwd=" & DBPassword
    DAO.DBEngine.CompactDatabase strSrcDir & "LabPatientValues.mdb", strSrcDir & "tLabPatientValues.mdb", dbLangGeneral, , ";pwd=" & DBPassword
    DAO.DBEngine.CompactDatabase strSrcDir & "InstrumentRValues.mdb", strSrcDir & "tInstrumentRValues.mdb", dbLangGeneral, , ";pwd=" & DBPassword
    fsoFiles.CopyFile strSrcDir & "tLaboratory.mdb", strSrcDir & "Laboratory.mdb", True
    fsoFiles.CopyFile strSrcDir & "tInstrumentResults.mdb", strSrcDir & "InstrumentResults.mdb", True
    fsoFiles.CopyFile strSrcDir & "tLabBuffer.mdb", strSrcDir & "LabBuffer.mdb", True
    fsoFiles.CopyFile strSrcDir & "tLabPatientData.mdb", strSrcDir & "LabPatientData.mdb", True
    fsoFiles.CopyFile strSrcDir & "tLabPatientValues.mdb", strSrcDir & "LabPatientValues.mdb", True
    fsoFiles.CopyFile strSrcDir & "tInstrumentRValues.mdb", strSrcDir & "InstrumentRValues.mdb", True
    fsoFiles.DeleteFile strSrcDir & "tLaboratory.mdb", True
    fsoFiles.DeleteFile strSrcDir & "tInstrumentResults.mdb", True
    fsoFiles.DeleteFile strSrcDir & "tLabBuffer.mdb", True
    fsoFiles.DeleteFile strSrcDir & "tLabPatientData.mdb", True
    fsoFiles.DeleteFile strSrcDir & "tLabPatientValues.mdb", True
    fsoFiles.DeleteFile strSrcDir & "tInstrumentRValues.mdb", True
    MsgBox "Backup Successful. Press Ok to continue.", vbInformation + vbOKOnly, App.ProductName
End If
GetDBSettings
OpenConnection
setHostForms
End Sub

Public Function showBackupList() As String
    strBackupPath = ""
    frmBackupList.Show 1
    showBackupList = strBackupPath
End Function

Private Function checkResultString() As Boolean
Dim rstTemp As ADODB.Recordset
On Error GoTo errFound
Set rstTemp = myConnection.Execute("select top 1 * from tbLabInstrumentResults")
If rstTemp.Fields("Result").Type = 202 Then
    checkResultString = True
Else
    checkResultString = False
End If
errFound:
If Err.Number = -2147467259 Then
    FixLink DBDatabase, False
    Resume
End If
End Function

Public Sub setFormColors(myForm As Form)
Dim ctr As Integer
On Error Resume Next
    myForm.BackColor = GetSetting("Lab-I", "FormColors", "BodyFrame", &H6B4B01)
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

Public Function GenerateResultID(myStringValue As String) As String
Dim strLen As Integer
Dim ctr As Integer
Dim strTemp As String
Debug.Print myStringValue
strLen = Len(myStringValue)
strTemp = ""
For ctr = 1 To strLen
    strTemp = strTemp & Hex$(Asc(Mid(myStringValue, ctr, 1)))
Next ctr
GenerateResultID = strTemp
End Function


Public Function WinVersion() As String
Dim OSInfo As OSVERSIONINFO
Dim retvalue As Integer
  
OSInfo.dwOSVersionInfoSize = 148
OSInfo.szCSDVersion = Space$(128)
retvalue = GetVersionExA(OSInfo)
  
With OSInfo
     Select Case .dwPlatformId
         Case VER_PLATFORM_WIN32s           ' Win32s on Windows 3.1
              WinVersion = "Windows 3.1"
               
         Case VER_PLATFORM_WIN32_WINDOWS    ' Windows 95, Windows 98,
              Select Case .dwMinorVersion   ' or Windows Me
                  Case 0
                      WinVersion = "Windows 95"
                  Case 10
                      If (OSInfo.dwBuildNumber And &HFFFF&) = 2222 Then
                          WinVersion = "Windows 98SE"
                      Else
                          WinVersion = "Windows 98"
                      End If
                  Case 90
                      WinVersion = "Windows Me"
              End Select
     
         Case VER_PLATFORM_WIN32_NT         ' Windows NT, Windows 2000, Windows XP,
              Select Case .dwMajorVersion   ' or Windows Server 2003 family.
                  Case 3
                      WinVersion = "Windows NT 3.51"
                  Case 4
                      WinVersion = "Windows NT 4.0"
                  Case 5
                      Select Case .dwMinorVersion
                          Case 0
                              WinVersion = "Windows 2000"
                          Case 1
                              WinVersion = "Windows XP"
                          Case 2
                              WinVersion = "Windows Server 2003"
                      End Select
                    Case 6
                      Select Case .dwMinorVersion
                          Case 0
                              WinVersion = "Windows Vista"
                          Case 1
                              WinVersion = "Windows 7"
                          Case 2
                              WinVersion = "Windows 8"
                          Case 3
                              WinVersion = "Windows 8.1"
                      End Select
                    Case 10
                        WinVersion = "Windows 10"
              End Select
         Case Else
              WinVersion = "Failed"
     End Select
             
End With
End Function

Public Function GetSpecialFolderPath(ByVal folder_number As Long) As String
Dim Path As String

    Path = Space$(MAX_PATH)
    If SHGetSpecialFolderPath(721092, Path, _
        folder_number, False) _
    Then
        GetSpecialFolderPath = Left$(Path, InStr(Path, Chr$(0)))
    Else
        GetSpecialFolderPath = "???"
    End If
End Function

Public Sub CopyToVistaFolder()
Dim strDir As String
Dim strCurDir As String
Dim strTgtDir As String
Dim strSaveDB As String
DoEvents
CloseConnection

strCurDir = ""
strTgtDir = ""

If UCase(Dir(App.Path & "\Reports", vbDirectory)) = "REPORTS" Then
    If Len(Dir(Environ$("AllUsersProfile") & "\Laboratory Instrument Interface", vbDirectory)) = 0 Then
        fsoFiles.Createfolder Environ$("AllUsersProfile") & "\Laboratory Instrument Interface"
        fsoFiles.Createfolder Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports"
        strCurDir = App.Path & "\Reports\"
    End If
    If Len(Dir(Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports", vbDirectory)) = 0 Then
        fsoFiles.Createfolder Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports"
        strCurDir = App.Path & "\Reports\"
    End If
ElseIf UCase(Dir(App.Path & "\AppData\Reports", vbDirectory)) = "REPORTS" Then
    If Len(Dir(Environ$("AllUsersProfile") & "\Laboratory Instrument Interface", vbDirectory)) = 0 Then
        fsoFiles.Createfolder Environ$("AllUsersProfile") & "\Laboratory Instrument Interface"
        fsoFiles.Createfolder Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports"
        strCurDir = App.Path & "\AppData\Reports\"
    End If
    If Len(Dir(Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports", vbDirectory)) = 0 Then
        fsoFiles.Createfolder Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports"
        strCurDir = App.Path & "\AppData\Reports\"
    End If
End If

If Len(strCurDir) > 0 Then
    strTgtDir = Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports\"
    fsoFiles.CopyFile strCurDir & "*.*", strTgtDir, True
End If

strCurDir = ""
strTgtDir = ""
strSaveDB = Format(Replace(Format(DBDatabase, ">"), "LABORATORY.MDB", ""), ">")

If UCase(Dir(App.Path & "\Data", vbDirectory)) = "DATA" Then
    If Len(Dir(Environ$("AllUsersProfile") & "\Laboratory Instrument Interface", vbDirectory)) = 0 Then
        fsoFiles.Createfolder Environ$("AllUsersProfile") & "\Laboratory Instrument Interface"
        fsoFiles.Createfolder Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Data"
        strCurDir = App.Path & "\Data\"
    End If
    If Len(Dir(Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Data", vbDirectory)) = 0 Then
        fsoFiles.Createfolder Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Data"
        strCurDir = App.Path & "\Data\"
    End If
ElseIf UCase(Dir(App.Path & "\AppData\Data", vbDirectory)) = "DATA" Then
    If Len(Dir(Environ$("AllUsersProfile") & "\Laboratory Instrument Interface", vbDirectory)) = 0 Then
        fsoFiles.Createfolder Environ$("AllUsersProfile") & "\Laboratory Instrument Interface"
        fsoFiles.Createfolder Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Data"
        strCurDir = App.Path & "\AppData\Data\"
    End If
    If Len(Dir(Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports", vbDirectory)) = 0 Then
        fsoFiles.Createfolder Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Data"
        strCurDir = App.Path & "\AppData\Data\"
    End If
End If

strCurDir = Format(strCurDir, ">")
If Len(strCurDir) > 0 Then
    If strCurDir <> strSaveDB Then
        strCurDir = strSaveDB
    End If
    strTgtDir = Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Data\"
    fsoFiles.CopyFile strCurDir & "Laboratory.mdb", strTgtDir & "Laboratory.mdb", True
    fsoFiles.CopyFile strCurDir & "InstrumentResults.mdb", strTgtDir & "InstrumentResults.mdb", True
    fsoFiles.CopyFile strCurDir & "LabBuffer.mdb", strTgtDir & "LabBuffer.mdb", True
    fsoFiles.CopyFile strCurDir & "LabPatientData.mdb", strTgtDir & "LabPatientData.mdb", True
    fsoFiles.CopyFile strCurDir & "LabPatientValues.mdb", strTgtDir & "LabPatientValues.mdb", True
    fsoFiles.CopyFile strCurDir & "InstrumentRValues.mdb", strTgtDir & "InstrumentRValues.mdb", True
    fsoFiles.CopyFile strCurDir & "LabRequestBuffer.mdb", strTgtDir & "LabRequestBuffer.mdb", True
    fsoFiles.CopyFile strCurDir & "LabRequestForm.mdb", strTgtDir & "LabRequestForm.mdb", True
    SaveSetting "LAB-i", "LOCAL", "DATABASE", Blowfish.EncryptString(strTgtDir & "Laboratory.mdb", myKey, True)
    SaveSetting "LAB-i", "LOCAL", "DATABASEPASSWORD", Blowfish.EncryptString("instantiation", myKey, True)
    GetDBSettings
    OpenConnection
    FixLink DBDatabase
End If

'strCurDir = ""
'strTgtDir = ""
'
'
'strCurDir = Replace(Format(DBDatabase, ">"), "LABORATORY.MDB", "")
'strTgtDir = Left(VistaPath, Len(VistaPath) - 1) & "\" & App.Title
'strDir = Dir(strTgtDir, vbDirectory)
'If Len(Trim(strDir)) = 0 Then
'    fsoFiles.Createfolder strTgtDir
'End If
'strTgtDir = strTgtDir & "\Data"
'strDir = Dir(strTgtDir, vbDirectory)
'If Len(Trim(strDir)) = 0 Then
'    fsoFiles.Createfolder strTgtDir
'End If
'strTgtDir = strTgtDir & "\"
'fsoFiles.CopyFile strCurDir & "Laboratory.mdb", strTgtDir & "Laboratory.mdb", True
'fsoFiles.CopyFile strCurDir & "InstrumentResults.mdb", strTgtDir & "InstrumentResults.mdb", True
'fsoFiles.CopyFile strCurDir & "LabBuffer.mdb", strTgtDir & "LabBuffer.mdb", True
'fsoFiles.CopyFile strCurDir & "LabPatientData.mdb", strTgtDir & "LabPatientData.mdb", True
'fsoFiles.CopyFile strCurDir & "LabPatientValues.mdb", strTgtDir & "LabPatientValues.mdb", True
'fsoFiles.CopyFile strCurDir & "InstrumentRValues.mdb", strTgtDir & "InstrumentRValues.mdb", True
'fsoFiles.CopyFile strCurDir & "LabRequestBuffer.mdb", strTgtDir & "LabRequestBuffer.mdb", True
'fsoFiles.CopyFile strCurDir & "LabRequestForm.mdb", strTgtDir & "LabRequestForm.mdb", True
'SaveSetting "LAB-i", "LOCAL", "DATABASE", Blowfish.EncryptString(strTgtDir & "Laboratory.mdb", myKey, True)
'SaveSetting "LAB-i", "LOCAL", "DATABASEPASSWORD", Blowfish.EncryptString("instantiation", myKey, True)
'GetDBSettings
'OpenConnection
'FixLink DBDatabase
End Sub

Function Mod43(C39 As String) As String
Dim i, total As Integer
  For i = 1 To Len(C39)
    total = InStr(charSet, Mid(C39, i, 1)) - 1 + total
  Next i
  Mod43 = C39 & Mid$(charSet, (total Mod 43 + 1), 1)
End Function

'--------------------------------------------
'HL7 Segments
'--------------------------------------------

Public Function HL7_MSH(mess As String) As String
Dim strTemp As String
Dim strDate As String
If InStr(1, mess, "DATE") > 0 Then
    strDate = Mid(mess, InStr(1, mess, "DATE") + 5, 8) & " " & Mid(mess, InStr(1, mess, "TIME") + 5, 8)
    If Not IsDate(strDate) Then strDate = CStr(Now)
Else
    strDate = CStr(Now)
End If
strDate = (DateDiff("d", Format(strDate, "mm/dd/yyyy"), "1/1/" & Format(strDate, "yyyy")) * -1) + 1 & "-" & Format(strDate, "YYHHNNss")
strTemp = ""
strTemp = "MSH" & "|" & GetSetting("LAB-i", "HL7 Settings", "Encoding Characters", "^~\&") & "|" & GetSetting("LAB-i", "HL7 Settings", "Receiving Application", "MDMPI") & _
    "||" & GetSetting("LAB-i", "HL7 Settings", "Sending Application", "LAB-Iv" & App.Major) & "||" & Format(Now, "YYYYMMDDHHnnss") & "||ORU^R01|" & strDate & "|P|" & GetSetting("LAB-i", "HL7 Settings", "HL7 Version", "2.5.1")
    HL7_MSH = strTemp
End Function

Public Function HL7_PID(mess As String, Optional HL7Request As String = "") As String
Dim strTemp As String
Dim tempPatientName, LastName, FirstName, bDate As String
Dim Temp() As String
strTemp = ""
tempPatientName = Trim(Format(getMessageValues(mess, "ID2"), ">"))
If InStr(1, tempPatientName, ",") > 0 Then
    Temp = Split(tempPatientName, ",")
    LastName = Trim(Temp(0))
    FirstName = Trim(Temp(1))
ElseIf InStr(1, tempPatientName, " ") > 0 Then
    Temp = Split(tempPatientName, " ")
    LastName = Trim(Temp(1))
    FirstName = Trim(Temp(0))
Else
    LastName = tempPatientName
    FirstName = ""
End If
bDate = Trim(Format(getMessageValues(mess, "BIRTH"), ">"))
If IsDate(bDate) Then bDate = Format(bDate, "YYYYMMDD") Else bDate = ""
strTemp = "PID|||" & Trim(Format(getMessageValues(mess, "ID1"), ">")) & "^^^^PIN||" & LastName & "^" & FirstName & "^^^^^^^^^^^" & _
"||" & bDate & "|" & Trim(Format(getMessageValues(mess, "SEX"), ">"))
If Len(Trim(HL7Request)) = 0 Then HL7_PID = strTemp Else HL7_PID = HL7Request
End Function

Public Function HL7_ORC(mess As String, Optional HL7Request As String = "") As String
Dim strTemp As String
Dim arrTemp() As String
Dim arrTemp2() As String
Dim ctr As Integer
If Len(Trim(HL7Request)) > 0 Then
    If InStr(1, HL7Request, vbCrLf) > 0 Then
        arrTemp2 = Split(HL7Request, vbCrLf)
        strTemp = ""
        For ctr = 0 To UBound(arrTemp2)
            arrTemp = Split(arrTemp2(ctr), "|")
            If ctr = 0 Then
                strTemp = "ORC|RE||" & arrTemp(3) & "|" & arrTemp(4) & "|" & "CM" & "||||" & Format(Now, "YYYYMMDDHHnnss") & "|" & "" & "|||||" & arrTemp(15)
            ElseIf ctr > 0 Then
                strTemp = strTemp & vbCrLf & "ORC|RE||" & arrTemp(3) & "|" & arrTemp(4) & "|" & "CM" & "||||" & Format(Now, "YYYYMMDDHHnnss") & "|" & "" & "|||||" & arrTemp(15)
            End If
        Next ctr
    Else
        arrTemp = Split(HL7Request, "|")
        strTemp = ""
        strTemp = "ORC|RE||" & arrTemp(3) & "|" & arrTemp(4) & "|" & "CM" & "||||" & Format(Now, "YYYYMMDDHHnnss") & "|" & "" & "|||||" & arrTemp(15)
    End If
Else
    strTemp = ""
    strTemp = "ORC|RE||" & Trim(Format(getMessageValues(mess, "SEQUENCE"), ">")) & "||CM||||" & Format(Now, "YYYYMMDDHHnnss") & "|" & "" & "|||||"
End If
HL7_ORC = strTemp
End Function

Public Function HL7_OBR(mess As String, Optional HL7Request As String = "") As String
Dim strTemp As String
Dim StartDate, EndDate As String
Dim strDate As String
Dim arrTemp() As String
Dim arrTemp2() As String
Dim ctr As Integer
If InStr(1, mess, "DATE") > 0 Then
    strDate = Mid(mess, InStr(1, mess, "DATE") + 5, 8) & " " & Mid(mess, InStr(1, mess, "TIME") + 5, 8)
    If Not IsDate(strDate) Then strDate = CStr(Now)
Else
    strDate = CStr(Now)
End If
strDate = (DateDiff("d", Format(strDate, "mm/dd/yyyy"), "1/1/" & Format(strDate, "yyyy")) * -1) + 1 & "-" & Format(strDate, "YYHHNNss")
If Len(Trim(HL7Request)) > 0 Then
    If InStr(1, HL7Request, vbCrLf) > 0 Then
        arrTemp2 = Split(HL7Request, vbCrLf)
        strTemp = ""
        For ctr = 0 To UBound(arrTemp2)
            arrTemp = Split(arrTemp2(ctr), "|")
            StartDate = Trim(Format(getMessageValues(mess, "ED"), ">")) & " " & Trim(Format(getMessageValues(mess, "ET"), ">"))
            EndDate = Trim(Format(getMessageValues(mess, "DATE"), ">")) & " " & Trim(Format(getMessageValues(mess, "TIME"), ">"))
            If IsDate(StartDate) Then StartDate = Format(StartDate, "YYYYMMDDHHnnss") Else StartDate = ""
            If IsDate(EndDate) Then EndDate = Format(EndDate, "YYYYMMDDHHnnss") Else EndDate = ""
            If ctr = 0 Then
                strTemp = "OBR|1|" & arrTemp(2) & "|" & arrTemp(3) & "|" & arrTemp(4) & "|||" & StartDate & "|" & EndDate & "||||||" & arrTemp(14) & "|||||||||||||" & arrTemp(27)
            Else
                strTemp = strTemp & vbCrLf & "OBR|" & ctr + 1 & "|" & arrTemp(2) & "|" & arrTemp(3) & "|" & arrTemp(4) & "|||" & StartDate & "|" & EndDate & "||||||" & arrTemp(14) & "|||||||||||||" & arrTemp(27)
            End If
        Next ctr
    Else
        arrTemp = Split(HL7Request, "|")
        strTemp = ""
        StartDate = Trim(Format(getMessageValues(mess, "ED"), ">")) & " " & Trim(Format(getMessageValues(mess, "ET"), ">"))
        EndDate = Trim(Format(getMessageValues(mess, "DATE"), ">")) & " " & Trim(Format(getMessageValues(mess, "TIME"), ">"))
        If IsDate(StartDate) Then StartDate = Format(StartDate, "YYYYMMDDHHnnss") Else StartDate = ""
        If IsDate(EndDate) Then EndDate = Format(EndDate, "YYYYMMDDHHnnss") Else EndDate = ""
        strTemp = "OBR|1|" & arrTemp(2) & "|" & arrTemp(3) & "|" & arrTemp(4) & "|||" & StartDate & "|" & EndDate & "||||||" & arrTemp(14) & "|||||||||||||" & arrTemp(27)
    End If
Else
    strTemp = ""
    StartDate = Trim(Format(getMessageValues(mess, "ED"), ">")) & " " & Trim(Format(getMessageValues(mess, "ET"), ">"))
    EndDate = Trim(Format(getMessageValues(mess, "DATE"), ">")) & " " & Trim(Format(getMessageValues(mess, "TIME"), ">"))
    If IsDate(StartDate) Then StartDate = Format(StartDate, "YYYYMMDDHHnnss") Else StartDate = ""
    If IsDate(EndDate) Then EndDate = Format(EndDate, "YYYYMMDDHHnnss") Else EndDate = ""
    strTemp = "OBR|1||" & strDate & "||||" & StartDate & "|" & EndDate & "||||||" & StartDate & "|||||||||||||^^^^^R"
End If
HL7_OBR = strTemp
End Function

Public Function HL7_OBX(mess As String, rstExams As ADODB.Recordset, Optional HL7Request As String = "") As String
Dim strTemp As String
Dim ctr As Integer
Dim ctr2 As Integer
Dim strEndString As String
Dim arrTemp() As String
Dim arrTemp2() As String
Dim arrTemp3() As String
Dim arrNValues() As String
Dim mValues As String
mValues = ""
Dim rstTemp As ADODB.Recordset
    strEndString = "|||F" & vbCrLf
    strTemp = ""
    If InStr(1, HL7Request, vbCrLf) > 0 Then
        arrTemp3 = Split(HL7Request, vbCrLf)
        For ctr2 = 0 To UBound(arrTemp3)
            If Len(Trim(arrTemp3(ctr2))) > 0 Then
                arrTemp = Split(arrTemp3(ctr2), "|")
                arrTemp2 = Split(arrTemp(4), "^")
                Set rstTemp = myConnection.Execute("select a.InstrumentID, a.ProfileName, b.ExamID, b.SortNum, c.InstExamName, c.DisplayName, c.Unit, c.HisExamCode, c.ExamNotes from tbLabProfiles a , tbLabProfileExams b, tbCoLabExam c where a.ProfileID = b.ProfileID " & _
                                "and a.InstrumentID = '" & InstrumentID & "' and a.InstrumentID = b.InstrumentID and c.InstrumentID = a.InstrumentID and c.ExamID = b.ExamID and a.ProfileName = '" & arrTemp2(0) & "' order by b.SortNum")
            End If
            If rstTemp.RecordCount > 0 Then
                rstTemp.MoveFirst
                For ctr = 0 To rstTemp.RecordCount - 1
                    If Len(Trim(rstTemp.Fields("ExamNotes").Value) & "") > 0 Then
                        If InStr(1, Trim(rstTemp.Fields("ExamNotes").Value), "|") Then
                            arrNValues = Split(Trim(rstTemp.Fields("ExamNotes").Value), "|")
                            If Trim(getStrResult(1, mess, "SEX")) = "M" Then
                                mValues = arrNValues(0)
                            ElseIf Trim(getStrResult(1, mess, "SEX")) = "F" Then
                                mValues = arrNValues(1)
                            Else
                                mValues = ""
                            End If
                        Else
                            mValues = Trim(rstTemp.Fields("ExamNotes").Value)
                        End If
                    Else
                        mValues = ""
                    End If
                    strTemp = strTemp & "OBX|" & ctr + 1 & "|ST|" & rstTemp.Fields("HisExamCode").Value & "^" & rstTemp.Fields("DisplayName").Value & "||" & getStrResult(1, mess, rstTemp.Fields("InstExamName").Value) & "|" & rstTemp.Fields("Unit").Value & "|" & mValues & "|" & Replace(getStrResult(2, mess, rstTemp.Fields("InstExamName").Value), Chr(160), "") & strEndString
                    If Not rstTemp.EOF Then rstTemp.MoveNext
                Next ctr
                strTemp = strTemp & "<SP>"
            Else
                If rstExams.State = 1 Then
                    If rstExams.RecordCount > 0 Then
                        rstExams.MoveFirst
                        For ctr = 0 To rstExams.RecordCount - 1
                            strTemp = strTemp & "OBX|" & ctr + 1 & "|ST|" & rstExams.Fields("HisExamCode").Value & "^" & rstExams.Fields("DisplayName").Value & "||" & getStrResult(1, mess, rstExams.Fields("InstExamName").Value) & "|" & rstExams.Fields("Unit").Value & "|" & mValues & "|" & Replace(getStrResult(2, mess, rstExams.Fields("InstExamName").Value), Chr(160), "") & strEndString
                            If Not rstExams.EOF Then rstExams.MoveNext
                        Next ctr
                        strTemp = strTemp & "<SP>"
                    End If
                End If
            End If
        Next ctr2
    Else
        If Len(Trim(HL7Request)) > 0 Then
            arrTemp = Split(HL7Request, "|")
            arrTemp2 = Split(arrTemp(4), "^")
            Set rstTemp = myConnection.Execute("select a.InstrumentID, a.ProfileName, b.ExamID, b.SortNum, c.InstExamName, c.DisplayName, c.Unit, c.HisExamCode, c.ExamNotes from tbLabProfiles a , tbLabProfileExams b, tbCoLabExam c where a.ProfileID = b.ProfileID " & _
                            "and a.InstrumentID = '" & InstrumentID & "' and a.InstrumentID = b.InstrumentID and c.InstrumentID = a.InstrumentID and c.ExamID = b.ExamID and a.ProfileName = '" & arrTemp2(0) & "' order by b.SortNum")
        End If
        If rstTemp.RecordCount > 0 Then
            rstTemp.MoveFirst
            For ctr = 0 To rstTemp.RecordCount - 1
                If Len(Trim(rstTemp.Fields("ExamNotes").Value) & "") > 0 Then
                    If InStr(1, Trim(rstTemp.Fields("ExamNotes").Value), "|") Then
                        arrNValues = Split(Trim(rstTemp.Fields("ExamNotes").Value), "|")
                        If Trim(getStrResult(1, mess, "SEX")) = "M" Then
                            mValues = arrNValues(0)
                        ElseIf Trim(getStrResult(1, mess, "SEX")) = "F" Then
                            mValues = arrNValues(1)
                        Else
                            mValues = ""
                        End If
                    Else
                        mValues = Trim(rstTemp.Fields("ExamNotes").Value)
                    End If
                Else
                    mValues = ""
                End If
                strTemp = strTemp & "OBX|" & ctr + 1 & "|ST|" & rstTemp.Fields("HisExamCode").Value & "^" & rstTemp.Fields("DisplayName").Value & "||" & getStrResult(1, mess, rstTemp.Fields("InstExamName").Value) & "|" & rstTemp.Fields("Unit").Value & "|" & mValues & "|" & Replace(getStrResult(2, mess, rstTemp.Fields("InstExamName").Value), Chr(160), "") & strEndString
                If Not rstTemp.EOF Then rstTemp.MoveNext
            Next ctr
        Else
            If rstExams.State = 1 Then
                If rstExams.RecordCount > 0 Then
                    rstExams.MoveFirst
                    For ctr = 0 To rstExams.RecordCount - 1
                        strTemp = strTemp & "OBX|" & ctr + 1 & "|ST|" & rstExams.Fields("HisExamCode").Value & "^" & rstExams.Fields("DisplayName").Value & "||" & getStrResult(1, mess, rstExams.Fields("InstExamName").Value) & "|" & rstExams.Fields("Unit").Value & "|" & mValues & "|" & Replace(getStrResult(2, mess, rstExams.Fields("InstExamName").Value), Chr(160), "") & strEndString
                        If Not rstExams.EOF Then rstExams.MoveNext
                    Next ctr
                End If
            End If
        End If
    End If
    HL7_OBX = strTemp
End Function

Public Function HL7_ORU(mess As String, Optional rstExams As ADODB.Recordset) As String
Dim strTemp As String
Dim arrTemp() As String
Dim myPID As String
Dim myORC As String
Dim myOBR As String
Dim arrORC() As String
Dim arrOBR() As String
Dim arrOBX() As String
Dim rstRequest As ADODB.Recordset
Dim rstTemp As ADODB.Recordset
Dim ctr As Integer
Dim mySampleID As String
Dim myCount As Integer
Dim tempArr() As String
    strTemp = ""
    myCount = 0
    If MessageFormat = "STKS" Then
        mySampleID = Trim(Format(getMessageValues(mess, "PID1"), ">"))
        If Len(Trim(mySampleID)) = 0 Then
            mySampleID = Trim(Format(getMessageValues(mess, "ID1"), ">"))
        End If
    Else
        tempArr = Split(MsgDef.segment_oBI(mess), "|")
        mySampleID = tempArr(6)
    End If
    Set rstRequest = myConnection.Execute("Select top 1 * from tblabrequest where messageAccNum ='" & mySampleID & "' and messagestatus = '2'")
    If rstRequest.RecordCount = 1 Then
        arrTemp = Split(rstRequest.Fields("MessageData").Value, vbCr)
        myConnection.Execute "Update tbLabRequest set messageStatus = '3' where messageid = " & rstRequest.Fields("messageID").Value
        myCount = UBound(arrTemp)
    End If
    myPID = ""
    myORC = ""
    myOBR = ""
    For ctr = 0 To myCount - 1
        If Left(arrTemp(ctr), 3) = "PID" Then
            myPID = arrTemp(ctr)
        ElseIf Left(arrTemp(ctr), 3) = "ORC" Then
            If Len(Trim(myORC)) > 0 Then
                myORC = myORC & vbCrLf & arrTemp(ctr)
            Else
                myORC = arrTemp(ctr)
            End If
        ElseIf Left(arrTemp(ctr), 3) = "OBR" Then
            If Len(Trim(myOBR)) > 0 Then
                myOBR = myOBR & vbCrLf & arrTemp(ctr)
            Else
                myOBR = arrTemp(ctr)
            End If
        End If
    Next ctr
    If MessageFormat = "STKS" Then
        If InStr(1, myORC, vbCrLf) > 0 Then
            arrORC = Split(HL7_ORC(mess, myORC), vbCrLf)
            arrOBR = Split(HL7_OBR(mess, myOBR), vbCrLf)
            arrOBX = Split(HL7_OBX(mess, rstExams, myOBR), "<SP>")
            strTemp = ""
            For ctr = 0 To UBound(arrORC)
                strTemp = strTemp & arrOBR(ctr) & vbCrLf & arrOBX(ctr)
            Next ctr
            strTemp = HL7_MSH(mess) & vbCrLf & HL7_PID(mess, myPID) & vbCrLf & arrORC(0) & vbCrLf & strTemp
        Else
            strTemp = HL7_MSH(mess) & vbCrLf & HL7_PID(mess, myPID) & vbCrLf & HL7_ORC(mess, myORC) & vbCrLf & HL7_OBR(mess, myOBR) & vbCrLf & HL7_OBX(mess, rstExams, myOBR)
        End If
    ElseIf MessageFormat = "ASTM" Then
        If InStr(1, myORC, vbCrLf) > 0 Then
            arrORC = Split(MsgDef.HL7_ORC(mess, myORC), vbCrLf)
            arrOBR = Split(MsgDef.HL7_OBR(mess, myOBR), vbCrLf)
            arrOBX = Split(MsgDef.HL7_OBX(mess, myOBR, myConnection), "<SP>")
            strTemp = ""
            For ctr = 0 To UBound(arrORC)
                strTemp = strTemp & arrOBR(ctr) & vbCrLf & arrOBX(ctr) & vbCrLf
            Next ctr
            strTemp = MsgDef.HL7_MSH(mess) & vbCrLf & MsgDef.HL7_PID(mess, myPID) & vbCrLf & arrORC(0) & vbCrLf & strTemp
        Else
            strTemp = MsgDef.HL7_MSH(mess) & vbCrLf & MsgDef.HL7_PID(mess, myPID) & vbCrLf & MsgDef.HL7_ORC(mess, myORC) & vbCrLf & MsgDef.HL7_OBR(mess, myOBR) & vbCrLf & MsgDef.HL7_OBX(mess, myOBR, myConnection)
        End If
    End If
    Debug.Print strTemp
    HL7_ORU = strTemp
End Function

Public Sub WriteToFile(mText As String, mDirectory As String)
Dim strLocation As String
Dim mFilename As String
Dim mFile
If Len(Trim(GetSetting("LAB-i", "HL7 Settings", "FWEXT", "txt"))) > 0 Then
    mFilename = Format(Now, "yyyyMMddHHnnss") & Right(Format(Timer, "#0.00"), 2) & "." & GetSetting("LAB-i", "HL7 Settings", "FWEXT", "txt")
Else
    mFilename = Format(Now, "yyyyMMddHHnnss") & Right(Format(Timer, "#0.00"), 2)
End If
strLocation = mDirectory & "\"
Set mFile = fsoFiles.CreateTextFile(strLocation & mFilename, False)
mFile.Write mText
mFile.Close
End Sub

Public Function SaveHL7toDB(mDirectory As String, mFileExt As String) As String
Dim mFilename As String
Dim intFileNumber As Integer
Dim myString As String
Dim myArray1() As String
Dim myArray2() As String
Dim myArray3() As String
Dim tempstr As String
Dim ctr As Integer
Dim myDate, myPIN, myName, myLName, myFname, myMName, mySampleID, myDOB, mySex As String
Dim mySampleIDOrg As String
Dim rstTemp As ADODB.Recordset
On Error GoTo ErrTrap
If Len(Trim(mFileExt)) = 0 Then
    mFilename = Dir(mDirectory & "\*")
Else
    mFilename = Dir(mDirectory & "\*." & mFileExt)
End If
SaveHL7toDB = mFilename
If Len(Trim(mFilename)) > 0 Then
    mFilename = mDirectory & "\" & mFilename
    intFileNumber = FreeFile
    Open mFilename For Input As #intFileNumber
    myString = Input(LOF(intFileNumber), #intFileNumber)
'-----------------------------
    myArray1 = Split(myString, vbCr)
    For ctr = 0 To UBound(myArray1) - 1
        myArray2 = Split(myArray1(ctr), "|")
        If Trim(myArray2(0)) = "MSH" Then
            myDate = myArray2(6)
        ElseIf Trim(myArray2(0)) = "PID" Then
            myArray3 = Split(myArray2(3), "^")
            myPIN = myArray3(0)
            myArray3 = Split(myArray2(5), "^")
            myLName = myArray3(0)
            myFname = myArray3(1)
            myMName = myArray3(2)
            myName = Trim(myLName & ", " & myFname & " " & myMName)
            myDOB = myArray2(7)
            mySex = myArray2(8)
        ElseIf Trim(myArray2(0)) = "ORC" Then
            mySampleID = myArray2(4)
            mySampleID = Right(mySampleID, 14)
            mySampleIDOrg = myArray2(4)
        End If
    Next ctr
    Set rstTemp = myConnection.Execute("Select * from tblabrequest where MessageAccNum = '" & mySampleID & "'")
    If rstTemp.RecordCount = 0 Then
        If Val(GetSetting("LAB-i", "HL7 Settings", "AutoSendRequest", 0)) = 1 Then
            myConnection.Execute "Insert into tblabrequest (messagename, messagelname, messagefname, messagemname, messagedate, messagepin, messageaccnum, messageaccnumorg, messagedata, messageDOB, messageSex, messagestatus) " & _
            "Select '" & Replace(myName, "'", "''") & "','" & Replace(myLName, "'", "''") & "','" & Replace(myFname, "'", "''") & "','" & Replace(myMName, "'", "''") & "','" & myDate & "','" & myPIN & "','" & mySampleID & "','" & mySampleIDOrg & "','" & Replace(myString, "'", "''") & "','" & myDOB & "','" & mySex & "','1'"
        Else
            myConnection.Execute "Insert into tblabrequest (messagename, messagelname, messagefname, messagemname, messagedate, messagepin, messageaccnum, messageaccnumorg, messagedata, messageDOB, messageSex, messagestatus) " & _
            "Select '" & Replace(myName, "'", "''") & "','" & Replace(myLName, "'", "''") & "','" & Replace(myFname, "'", "''") & "','" & Replace(myMName, "'", "''") & "','" & myDate & "','" & myPIN & "','" & mySampleID & "','" & mySampleIDOrg & "','" & Replace(myString, "'", "''") & "','" & myDOB & "','" & mySex & "','0'"
        End If
    End If
'-----------------------------
    Close #intFileNumber
    Kill mFilename
End If
ErrTrap:
    If Err.Number <> 0 Then
        SaveHL7toDB = Err.Description
        Close #intFileNumber
        Kill mFilename
    End If
End Function

Public Function ProcessHL7Request(myMessage As String, Optional myRack As String, Optional myCass As String, Optional myTrans As String = "N") As String
Dim myArray1() As String
Dim myArray2() As String
Dim myArray3() As String
Dim tempstr As String
Dim tempStr2 As String
Dim ctr As Integer
Dim ctr2 As Integer
Dim ctr3 As Integer
Dim myDate, myPIN, myName, myLName, myFname, myMName, mySampleID, myDOB, mySex As String
Dim rstTemp As ADODB.Recordset
Dim myTests As String
Dim arrTests() As String
Dim arrTests2() As String
Dim multiTest As String
multiTest = ""
ctr2 = 1
myTests = ""
myMessage = Replace(myMessage, vbLf, "")
myArray1 = Split(myMessage, vbCr)
For ctr = 0 To UBound(myArray1)
    If Len(Trim(myArray1(ctr))) > 0 Then
        myArray2 = Split(myArray1(ctr), "|")
        If Trim(myArray2(0)) = "MSH" Then
            myDate = myArray2(6)
        ElseIf Trim(myArray2(0)) = "PID" Then
            myArray3 = Split(myArray2(3), "^")
            myPIN = myArray3(0)
            myArray3 = Split(myArray2(5), "^")
            myLName = myArray3(0)
            myFname = myArray3(1)
            myMName = myArray3(2)
            myName = Trim(myLName & ", " & myFname & " " & myMName)
            myDOB = myArray2(7)
            mySex = myArray2(8)
        ElseIf Trim(myArray2(0)) = "ORC" Then
            mySampleID = myArray2(4)
        ElseIf Trim(myArray2(0)) = "OBR" Then
            If Len(Trim(myTests)) > 0 Then
                myTests = myTests & vbCrLf & myArray2(4)
            Else
                myTests = myArray2(4)
            End If
        End If
    End If
Next ctr
If MessageFormat = "STKS" Then
    tempstr = Chr(1) & "07" & vbCrLf _
              & "WLAD" & vbCrLf _
              & "ID " & myPIN & vbCrLf _
              & "PL " & myLName & vbCrLf _
              & "PF " & myFname & vbCrLf _
              & "BD " & Mid(myDOB, 5, 2) & "/" & Mid(myDOB, 7, 2) & "/" & Mid(myDOB, 1, 4) & vbCrLf _
              & "SX " & mySex & vbCrLf _
              & "TS DIFF," & mySampleID & "," & vbCrLf & vbCrLf _
              & "-------------" & vbCrLf
    tempstr = tempstr & String(256 - Len(tempstr), " ")
    tempstr = "01" & Chr(2) & "01" & tempstr & CRC16_CCITT_M(tempstr) & Chr(3)
ElseIf MessageFormat = "ASTM" Then
    tempStr2 = ctr2 & "H|\^&"
    ctr2 = ctr2 + 1
    tempstr = Chr(2) & tempStr2 & vbCr & Chr(3) & ASTMChecksum(tempStr2) & vbCrLf
    tempStr2 = ctr2 & "P|1||" & myPIN & "||" & myLName & "^" & myFname & "^" & myMName & "||" & myDOB & "^^|" & mySex & "|||||||||||||||||"
    ctr2 = ctr2 + 1
    tempstr = tempstr & Chr(2) & tempStr2 & vbCr & Chr(3) & ASTMChecksum(tempStr2) & vbCrLf
'    If InStr(1, myTests, vbCrLf) Then
        arrTests = Split(myTests, vbCrLf)
        ctr3 = 1
        For ctr = 0 To UBound(arrTests)
            arrTests2 = Split(arrTests(ctr), "^")
            Set rstTemp = myConnection.Execute("select a.profilename, b.examid, b.group, c.InstExamName from " & _
                "tblabprofiles a, tblabprofileexams b, tbcolabexam c where " & _
                "a.Instrumentid = '" & InstrumentID & "' and a.instrumentid = b.Instrumentid " & _
                "and a.Instrumentid = c.InstrumentID and a.profileid= b.profileid and b.ExamID = c.ExamID " & _
                "and a.ProfileName ='" & arrTests2(0) & "'")
            If rstTemp.RecordCount > 1 Then
                rstTemp.MoveFirst
                Do While Not rstTemp.EOF
                    If Len(Trim(multiTest)) = 0 Then
                        multiTest = "^^^" & rstTemp("InstExamName").Value & "^^N"
                    Else
                        multiTest = multiTest & "\" & "^^^" & rstTemp("InstExamName").Value & "^^N"
                    End If

'                    tempStr2 = ctr2 & "O|" & ctr3 & "|" & Right(mySampleID, Len(mySampleID) - 2) & "^^^^||^^^" & rstTemp("InstExamName").Value & "^^N|R||" & myDate & "||||N||||S|||1^1.00"
'                    tempStr = tempStr & Chr(2) & tempStr2 & vbCr & Chr(3) & ASTMChecksum(tempStr2) & vbCrLf
'                    If ctr2 = 7 Then ctr2 = 0 Else ctr2 = ctr2 + 1
'                    ctr3 = ctr3 + 1
                    rstTemp.MoveNext
                Loop
'                tempStr2 = ctr2 & "O|" & ctr3 & "|" & Right(mySampleID, Len(mySampleID) - 2) & "^^^^||" & multiTest & "|R||" & myDate & "||||N||||S|||1^1.00"
'                tempStr = tempStr & Chr(2) & tempStr2 & vbCr & Chr(3) & ASTMChecksum(tempStr2) & vbCrLf
'                If ctr2 = 7 Then ctr2 = 0 Else ctr2 = ctr2 + 1
'                ctr3 = ctr3 + 1
            ElseIf rstTemp.RecordCount = 1 Then
                    If Len(Trim(multiTest)) = 0 Then
                        multiTest = "^^^" & rstTemp("InstExamName").Value & "^^N"
                    Else
                        multiTest = multiTest & "\" & "^^^" & rstTemp("InstExamName").Value & "^^N"
                    End If
'                    tempStr2 = ctr2 & "O|" & ctr3 & "|" & Right(mySampleID, Len(mySampleID) - 2) & "^^^^||^^^" & rstTemp("InstExamName").Value & "^^N|R||" & myDate & "||||N||||S|||1^1.00"
'                    tempStr = tempStr & Chr(2) & tempStr2 & vbCr & Chr(3) & ASTMChecksum(tempStr2) & vbCrLf
'                    If ctr2 = 7 Then ctr2 = 0 Else ctr2 = ctr2 + 1
'                    ctr3 = ctr3 + 1
            End If
        Next ctr
        tempStr2 = ctr2 & "O|" & ctr3 & "|" & mySampleID & "^" & myRack & "^" & myCass & "^^||" & multiTest & "|R||" & myDate & "||||" & myTrans & "||||S|||1^1.00"
        tempstr = tempstr & Chr(2) & tempStr2 & vbCr & Chr(3) & ASTMChecksum(tempStr2) & vbCrLf
        If ctr2 = 7 Then ctr2 = 0 Else ctr2 = ctr2 + 1
        ctr3 = ctr3 + 1

        ' rstTemp.Close
    tempStr2 = ctr2 & "L|1|N"
    tempstr = tempstr & Chr(2) & tempStr2 & vbCr & Chr(3) & ASTMChecksum(tempStr2) & vbCrLf & Chr(4)
'    Debug.Print tempStr
End If
ProcessHL7Request = tempstr
End Function

Public Function isTransparent(ByVal hWnd As Long) As Boolean
On Error Resume Next
Dim Msg As Long
Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
  isTransparent = True
Else
  isTransparent = False
End If
If Err Then
  isTransparent = False
End If
End Function

Public Function MakeTransparent(ByVal hWnd As Long, Perc As Integer) As Long
Dim Msg As Long
On Error Resume Next
If Perc < 0 Or Perc > 255 Then
  MakeTransparent = 1
Else
  Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
  Msg = Msg Or WS_EX_LAYERED
  SetWindowLong hWnd, GWL_EXSTYLE, Msg
  SetLayeredWindowAttributes hWnd, 0, Perc, LWA_ALPHA
  MakeTransparent = 0
End If
If Err Then
  MakeTransparent = 2
End If
End Function

Public Function MakeOpaque(ByVal hWnd As Long) As Long
Dim Msg As Long
On Error Resume Next
Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
Msg = Msg And Not WS_EX_LAYERED
SetWindowLong hWnd, GWL_EXSTYLE, Msg
SetLayeredWindowAttributes hWnd, 0, 0, LWA_ALPHA
MakeOpaque = 0
If Err Then
  MakeOpaque = 2
End If
End Function


Private Sub VerifyRequestDB(mDBLocation As String)
Dim myStream() As Byte
Dim f As Long
Dim Location As String

If Right(mDBLocation, 1) = "\" Then
    Location = Left(mDBLocation, Len(mDBLocation) - 1)
Else
    Location = mDBLocation
End If

If Len(Dir(Location & "\LabRequestBuffer.mdb")) = 0 Then
    myStream = LoadResData("LABREQUESTBUFFER.MDB", "CUSTOM")
    f = FreeFile()
    Open Location & "\LabRequestBuffer.mdb" For Binary As #f
    Put #f, , myStream
    Close #f
End If

If Len(Dir(Location & "\LabRequestForm.mdb")) = 0 Then
    myStream = LoadResData("LABREQUESTFORM.MDB", "CUSTOM")
    f = FreeFile()
    Open Location & "\LabRequestForm.mdb" For Binary As #f
    Put #f, , myStream
    Close #f
End If
End Sub

Public Function SystemSerialNumber() As String
Dim mother_boards As Variant
Dim board As Variant
Dim wmi As Variant
Dim serial_numbers As String

    ' Get the Windows Management Instrumentation object.
    Set wmi = GetObject("WinMgmts:")

    ' Get the "base boards" (mother boards).
    Set mother_boards = wmi.InstancesOf("Win32_BaseBoard")
    For Each board In mother_boards
        serial_numbers = serial_numbers & ", " & _
            board.SerialNumber
    Next board
    If Len(serial_numbers) > 0 Then serial_numbers = _
        Mid$(serial_numbers, 3)

    SystemSerialNumber = serial_numbers
'    SystemSerialNumber = "8AR011"
End Function

Public Function GetSerialNumber( _
    ByVal sDrive As String) As Long

    If Len(sDrive) Then
        If InStr(sDrive, "\\") = 1 Then
            ' Make sure we end in backslash for UNC
            If Right$(sDrive, 1) <> "\" Then
                sDrive = sDrive & "\"
            End If
        Else
            ' If not UNC, take first letter as drive
            sDrive = Left$(sDrive, 1) & ":\"
        End If
    Else
        ' Else just use current drive
        sDrive = vbNullString
    End If

    ' Grab S/N -- Most params can be NULL
    Call GetVolumeInformation( _
        sDrive, vbNullString, 0, GetSerialNumber, _
        ByVal 0&, ByVal 0&, vbNullString, 0)
'    GetSerialNumber = "1629277180"
End Function

Public Function GetCpuID() As String
  Dim wmi, cpu, cpuid
  Set wmi = GetObject("winmgmts:")
  For Each cpu In wmi.InstancesOf("Win32_Processor")
   cpuid = cpuid + cpu.ProcessorID
  Next
  GetCpuID = cpuid
End Function

Public Function IsRegistered() As Boolean
Dim mDirectory As String
Dim mFile As String
Dim tempBool As Boolean
Dim iFileNo As Integer
Dim tempstr As String
Dim newStr As String
Dim arrTemp() As String
Dim ArrLicense() As String
Dim arrExpireDate() As String
Dim arrRegCode() As String
Dim HWID1 As String
Dim HWID2 As String
Dim HWID3 As String
Dim ctr As Integer
Dim ActiveExpireDate As String
Dim LicenseElapse As Integer
On Error GoTo errFound
iFileNo = FreeFile
tempBool = False
mDirectory = Replace(Format(DBDatabase, ">"), "LABORATORY.MDB", "")
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
        arrTemp = Split(newStr, "|")
        If arrTemp(0) = "HOSTLICENSE" Then
            SaveSetting "LAB-i", "Registration", "HOSTUID", Replace(Replace(arrTemp(1), "{", ""), "}", "")
            SaveSetting "LAB-i", "Registration", "HOSTWID1", arrTemp(2)
            SaveSetting "LAB-i", "Registration", "HOSTWID2", arrTemp(3)
            SaveSetting "LAB-i", "Registration", "HOSTWID3", arrTemp(4)
            HWID1 = Replace(GetCpuID, "-", "")
            HWID2 = Replace(GetSerialNumber(Left(App.Path, 3)), "-", "")
            HWID3 = Replace(SystemSerialNumber, "-", "")
            ActiveExpireDate = ""
            LicenseElapse = -1
            If HWID1 = arrTemp(2) And HWID2 = arrTemp(3) And HWID3 = arrTemp(4) Then
                ArrLicense = Split(arrTemp(7), "<SP>")
                For ctr = 0 To UBound(ArrLicense)
                    arrExpireDate = Split(ArrLicense(ctr), "^")
                    If DateDiff("d", Now, arrExpireDate(1)) > 0 Then
                        LicenseElapse = DateDiff("d", Now, arrExpireDate(1))
                        If ActiveExpireDate = "" Then
                            ActiveExpireDate = arrExpireDate(1)
                        Else
                            If LicenseElapse < DateDiff("d", Now, ActiveExpireDate) Then
                                ActiveExpireDate = arrExpireDate(1)
                            End If
                        End If
                    End If
                Next ctr
            End If
        End If
    End If
End If
If LicenseElapse <= 0 Then
    MsgBox "The Laboratory Instrument Interface System has expired." & vbCrLf & vbCrLf & _
        "Please contact Marsman Drysdale Medical Products, Inc. as soon as possible." & vbCrLf & vbCrLf & "Internet connection is required to retrieve the license.", vbOKOnly + vbInformation, App.ProductName
End If
If IsDate(ActiveExpireDate) And LicenseElapse > 0 Then
    tempBool = True
    SaveSetting "Lab-I", "Registration", "Expiration", ActiveExpireDate
End If
errFound:
    If Err.Number <> 0 Then
        IsRegistered = False
    Else
        IsRegistered = tempBool
    End If
End Function

Public Sub DailyBackup()
Dim myStream() As Byte
Dim f As Long
Dim Location As String
Dim mDBLocation As String
Dim strCommand As String
Dim oShell As Object
Dim strOutput As String
Dim TempFolder As String
Dim lngFileCount As Long
Dim StrFileName As String
Dim FilesToDelete As Long
Dim ctr As Long
Dim FileNameToDelete As String
mDBLocation = Replace(Format(DBDatabase, ">"), "LABORATORY.MDB", "")
If Right(mDBLocation, 1) = "\" Then
    Location = Left(mDBLocation, Len(mDBLocation) - 1)
    mDBLocation = Left(mDBLocation, Len(mDBLocation) - 1)
Else
    Location = mDBLocation
End If
If Dir$(Location & "\DailyBackUp", vbDirectory) = "" Then
    fsoFiles.Createfolder Location & "\DailyBackUp"
End If
Location = Location & "\DailyBackUp"
If Dir$(Location & "\~Temp", vbDirectory) = "" Then
    fsoFiles.Createfolder Location & "\~Temp"
End If
TempFolder = Location & "\~Temp"
fsoFiles.CopyFile mDBLocation & "\Laboratory.mdb", TempFolder & "\Laboratory.mdb"

If Len(Dir$(Location & "\7Z.EXE")) = 0 Then
    myStream = LoadResData("7Z.EXE", "CUSTOM")
    f = FreeFile()
    Open Location & "\7z.exe" For Binary As #f
    Put #f, , myStream
    Close #f
End If
If Len(Dir$(Location & "\7Z.DLL")) = 0 Then
    myStream = LoadResData("7Z.DLL", "CUSTOM")
    f = FreeFile()
    Open Location & "\7z.dll" For Binary As #f
    Put #f, , myStream
    Close #f
End If
If Len(Trim(Dir$(Location & "\" & Format(Now, "yyyymmdd") & ".zip"))) > 0 Then
    fsoFiles.DeleteFile Location & "\" & Format(Now, "yyyymmdd") & ".zip"
End If
strCommand = """" & Location & "\7z.exe"" a -tzip """ & Location & "\" & Format(Now, "yyyymmdd") & ".zip"" """ & TempFolder & "\Laboratory.mdb"" """ & mDBLocation & "\*.mdb"" -x!""" & mDBLocation & "\Laboratory.mdb"""
Set oShell = CreateObject("WSCript.shell")
oShell.run strCommand, 0, True
fsoFiles.DeleteFile Location & "\7z.exe"
fsoFiles.DeleteFile Location & "\7z.dll"
fsoFiles.DeleteFile TempFolder & "\*.*"
fsoFiles.DeleteFolder TempFolder
StrFileName = Dir$(Location & "\*.zip")
lngFileCount = 0
Do While Len(StrFileName) <> 0
    lngFileCount = lngFileCount + 1
    StrFileName = Dir$
Loop
If lngFileCount > 15 Then
    FilesToDelete = lngFileCount - 15
    StrFileName = Dir$(Location & "\*.zip")
    For ctr = 1 To FilesToDelete
        Do While Len(StrFileName) <> 0
            If Len(FilenameToDate(StrFileName)) = 0 Then
                FileNameToDelete = StrFileName
                Exit Do
            Else
                If Len(FileNameToDelete) = 0 Then
                    FileNameToDelete = StrFileName
                Else
                    If DateDiff("d", FilenameToDate(FileNameToDelete), Now) < DateDiff("d", FilenameToDate(StrFileName), Now) Then
                        FileNameToDelete = StrFileName
                    End If
                End If
            End If
            StrFileName = Dir$
        Loop
        fsoFiles.DeleteFile Location & "\" & FileNameToDelete
        FileNameToDelete = ""
        StrFileName = Dir$(Location & "\*.zip")
    Next ctr
End If
End Sub

Private Function FilenameToDate(mFilename As String) As String
Dim tempstr As String
tempstr = Mid(mFilename, 1, InStr(1, mFilename, ".") - 1)
If Len(Trim(tempstr)) = 8 Then
    tempstr = Mid(tempstr, 5, 2) & "/" & Mid(tempstr, 7, 2) & "/" & Mid(tempstr, 1, 4)
End If
If IsDate(tempstr) Then
    FilenameToDate = tempstr
Else
    FilenameToDate = ""
End If
End Function

Public Function GetTotalRAM() As String
Dim udtMemStatEx As MEMORYSTATUSEX
       udtMemStatEx.dwLength = Len(udtMemStatEx)
       Call GlobalMemoryStatusEx(udtMemStatEx)
'USAGE:
'Total Physical Memory
  GetTotalRAM = NumberInKB(CLargeInt(udtMemStatEx.ulTotalPhys.LoPart, udtMemStatEx.ulTotalPhys.HiPart)) '& " (" & Round(CLargeInt(udtMemStatEx.ulAvailPhys.LoPart, udtMemStatEx.ulAvailPhys.HiPart) / (CLargeInt(udtMemStatEx.ulTotalPhys.LoPart, udtMemStatEx.ulTotalPhys.HiPart)) * 100) & "% Free)"
''Available Physical Memory
'   GetTotalRAM = NumberInKB(CLargeInt(udtMemStatEx.ulAvailPhys.LoPart, udtMemStatEx.ulAvailPhys.HiPart))
''Total Virtual Memory
'   GetTotalRAM = NumberInKB(CLargeInt(udtMemStatEx.ulTotalVirtual.LoPart, udtMemStatEx.ulTotalVirtual.HiPart))
''Available Virtual Memory
'   GetTotalRAM = vbTab + NumberInKB(CLargeInt(udtMemStatEx.ulAvailVirtual.LoPart, udtMemStatEx.ulAvailVirtual.HiPart))
''Memory Load
'   GetTotalRAM = CStr(udtMemStatEx.dwMemoryLoad) + "%"
End Function

Public Function GetInterfaceVersion() As String
    GetInterfaceVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

Public Function GetSAVersion() As String
Dim tempObj As Object
    Set tempObj = CreateObject("HostForms_SA.ClassForms")
    GetSAVersion = tempObj.Version
    Set tempObj = Nothing
End Function

Public Sub SaveRegToFile(ByVal hKey As Long, ByVal sKeyName As String, _
    ByVal sRegFile As String, Optional ByVal bIncludeSubKeys As Boolean = True, _
    Optional ByVal bAppendToFile As Boolean = False)
    
    Dim handle As Integer
    Dim sFirstKeyPart As String
    Dim col As New Collection
    Dim regItem As Variant
    Dim sText As String
    Dim sQuote As String
    Dim sTemp As String
    Dim sHex As String
    Dim i As Long
    Dim vValue As Variant
    Dim iPointer As MousePointerConstants
    Dim sValueName As String
    
    sQuote = Chr$(34)
    
    On Error Resume Next
     
    'conver the hKey value to the descriptive string
    Select Case hKey
        Case HKEY_CLASSES_ROOT: sFirstKeyPart = "HKEY_CLASSES_ROOT\"
        Case HKEY_CURRENT_CONFIG: sFirstKeyPart = "HKEY_CURRENT_CONFIG\"
        Case HKEY_CURRENT_USER: sFirstKeyPart = "HKEY_CURRENT_USER\"
        Case HKEY_LOCAL_MACHINE: sFirstKeyPart = "HKEY_LOCAL_MACHINE\"
        Case HKEY_USERS: sFirstKeyPart = "HKEY_USERS\"
    End Select
    
    'this can be a long operation
    iPointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    'if the text won't be appended, add the "REGEDIT4" header
    If bAppendToFile = False Then
        sText = "REGEDIT4" & vbCrLf & vbCrLf
    Else
        'add the same header if the text will be appended to an
        'existent file that does not contain the header.
        ' This works only if the file exists but is empty.
        handle = FreeFile
        Open sRegFile For Binary As #handle
        ' read the string and close the file
        sTemp = Space$(LOF(handle))
        Get #handle, , sTemp
        Close #handle
        'if not found, add it
        If InStr(1, sTemp, "REGEDIT4") = 0 Then
            sText = "REGEDIT4" & vbCrLf & vbCrLf
        End If
    End If
    
    'save the key name with the format [keyname]
    sText = sText & "[" & sFirstKeyPart & sKeyName & "]" & vbCrLf
    
    'get the collection with all the values under this key
    Set col = EnumRegistryValuesEx(hKey, sKeyName)
    For Each regItem In col
        vValue = regItem(1)
        Select Case regItem(2)
            Case vbString
                'if the value is a string, check if it's a path by looking if
                ' the 3 characters
                'are in the form X:\. If so, replace a single "\" with "\\"
                If Left$(vValue, 3) Like "[A-Z,a-z]:\" Then vValue = Replace _
                    (vValue, "\", "\\")
                'quote it
                sTemp = sQuote & vValue & sQuote
            Case vbLong
                'if it's a long, save it with the format dword:num
                sTemp = "dword:" & CLng(vValue)
            Case vbArray + vbByte
                'if it's an array of bytes, save it with the format hex:num1,
                ' num2,num3,...
                sTemp = "hex:"
                For i = 0 To UBound(vValue)
                    sHex = Hex$(vValue(i))
                    'convert from long to hex
                    If Len(sHex) < 2 Then sHex = "0" & sHex
                    sTemp = sTemp & sHex & ","
                Next
                'remove the last comma
                sTemp = Left$(sTemp, Len(sTemp) - 1)
            Case Else
                sTemp = ""
        End Select
        'get the value name: if the string is empty, take @,
        '  else take that name and quote it
        sValueName = IIf(Len(regItem(0)) > 0, sQuote & regItem(0) & sQuote, "@")
        'save this line to the temporary text that will be saved
        sText = sText & sValueName & "=" & sTemp & vbCrLf
    Next
    sText = sText & vbCrLf
    
    handle = FreeFile
    'open the target file with Append or Output mode,
    '  according to the bAppendToFile parameter
    If bAppendToFile Then
        Open sRegFile For Append As #handle
    Else
        Open sRegFile For Output As #handle
    End If
    'save the text
    Print #handle, sText;
    Close #handle
    
    'call recursively this routine to save all the subkeys,
    '  if the bIncludeSubKeys param is true
    If bIncludeSubKeys Then
        Set col = EnumRegistryKeys(hKey, sKeyName)
        For Each regItem In col
            'note: the text will be added to the file just created for the
            'values in the root key
            SaveRegToFile hKey, sKeyName & "\" & regItem, sRegFile, True, True
        Next
    End If
    
    Screen.MousePointer = iPointer
    
End Sub

Public Function EnumRegistryValuesEx(ByVal hKey As Long, ByVal KeyName As String) As _
    Collection
    Dim handle As Long
    Dim Index As Long
    Dim valueType As Long
    Dim name As String
    Dim nameLen As Long
    Dim resLong As Long
    Dim resString As String
    Dim dataLen As Long
    Dim valueInfo(0 To 2) As Variant
    Dim retVal As Long
    
    ' initialize the result
    Set EnumRegistryValuesEx = New Collection
    
    ' Open the key, exit if not found.
    If Len(KeyName) Then
        If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
        ' in all cases, subsequent functions use hKey
        hKey = handle
    End If
    
    Do
        ' this is the max length for a key name
        nameLen = 260
        name = Space$(nameLen)
        ' prepare the receiving buffer for the value
        dataLen = 4096
        ReDim resBinary(0 To dataLen - 1) As Byte
        
        ' read the value's name and data
        ' exit the loop if not found
        retVal = RegEnumValue(hKey, Index, name, nameLen, ByVal 0&, valueType, _
            resBinary(0), dataLen)
        
        ' enlarge the buffer if you need more space
        If retVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To dataLen - 1) As Byte
            retVal = RegEnumValue(hKey, Index, name, nameLen, ByVal 0&, _
                valueType, resBinary(0), dataLen)
        End If
        ' exit the loop if any other error (typically, no more values)
        If retVal Then Exit Do
        
        ' retrieve the value's name
        valueInfo(0) = Left$(name, nameLen)
        
        ' return a value corresponding to the value type
        Select Case valueType
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                valueInfo(1) = resLong
                valueInfo(2) = vbLong
            Case REG_SZ, REG_EXPAND_SZ
                ' copy everything but the trailing null char
                resString = Space$(dataLen - 1)
                CopyMemory ByVal resString, resBinary(0), dataLen - 1
                valueInfo(1) = resString
                valueInfo(2) = vbString
            Case REG_BINARY
                ' shrink the buffer if necessary
                If dataLen < UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To dataLen - 1) As Byte
                End If
                valueInfo(1) = resBinary()
                valueInfo(2) = vbArray + vbByte
            Case REG_MULTI_SZ
                ' copy everything but the 2 trailing null chars
                resString = Space$(dataLen - 2)
                CopyMemory ByVal resString, resBinary(0), dataLen - 2
                valueInfo(1) = resString
                valueInfo(2) = vbString
            Case Else
                ' Unsupported value type - do nothing
        End Select
        
        ' add the array to the result collection
        ' the element's key is the value's name
        EnumRegistryValuesEx.Add valueInfo, valueInfo(0)
        
        Index = Index + 1
    Loop
   
    ' Close the key, if it was actually opened
    If handle Then RegCloseKey handle
        
End Function

Public Function EnumRegistryValues(ByVal hKey As Long, ByVal KeyName As String) As _
    Collection
    Dim handle As Long
    Dim Index As Long
    Dim valueType As Long
    Dim name As String
    Dim nameLen As Long
    Dim resLong As Long
    Dim resString As String
    Dim dataLen As Long
    Dim valueInfo(0 To 1) As Variant
    Dim retVal As Long
    
    ' initialize the result
    Set EnumRegistryValues = New Collection
    
    ' Open the key, exit if not found.
    If Len(KeyName) Then
        If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
        ' in all cases, subsequent functions use hKey
        hKey = handle
    End If
    
    Do
        ' this is the max length for a key name
        nameLen = 260
        name = Space$(nameLen)
        ' prepare the receiving buffer for the value
        dataLen = 4096
        ReDim resBinary(0 To dataLen - 1) As Byte
        
        ' read the value's name and data
        ' exit the loop if not found
        retVal = RegEnumValue(hKey, Index, name, nameLen, ByVal 0&, valueType, _
            resBinary(0), dataLen)
        
        ' enlarge the buffer if you need more space
        If retVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To dataLen - 1) As Byte
            retVal = RegEnumValue(hKey, Index, name, nameLen, ByVal 0&, _
                valueType, resBinary(0), dataLen)
        End If
        ' exit the loop if any other error (typically, no more values)
        If retVal Then Exit Do
        
        ' retrieve the value's name
        valueInfo(0) = Left$(name, nameLen)
        
        ' return a value corresponding to the value type
        Select Case valueType
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                valueInfo(1) = resLong
            Case REG_SZ, REG_EXPAND_SZ
                ' copy everything but the trailing null char
                resString = Space$(dataLen - 1)
                CopyMemory ByVal resString, resBinary(0), dataLen - 1
                valueInfo(1) = resString
            Case REG_BINARY
                ' shrink the buffer if necessary
                If dataLen < UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To dataLen - 1) As Byte
                End If
                valueInfo(1) = resBinary()
            Case REG_MULTI_SZ
                ' copy everything but the 2 trailing null chars
                resString = Space$(dataLen - 2)
                CopyMemory ByVal resString, resBinary(0), dataLen - 2
                valueInfo(1) = resString
            Case Else
                ' Unsupported value type - do nothing
        End Select
        
        ' add the array to the result collection
        ' the element's key is the value's name
        EnumRegistryValues.Add valueInfo, valueInfo(0)
        
        Index = Index + 1
    Loop
   
    ' Close the key, if it was actually opened
    If handle Then RegCloseKey handle
        
End Function

Public Function EnumRegistryKeys(ByVal hKey As Long, ByVal KeyName As String) As _
    Collection
    Dim handle As Long
    Dim Length As Long
    Dim Index As Long
    Dim subkeyName As String
    
    ' initialize the result collection
    Set EnumRegistryKeys = New Collection
    
    ' Open the key, exit if not found
    If Len(KeyName) Then
        If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
        ' in all case the subsequent functions use hKey
        hKey = handle
    End If
    
    Do
        ' this is the max length for a key name
        Length = 260
        subkeyName = Space$(Length)
        ' get the N-th key, exit the loop if not found
        If RegEnumKey(hKey, Index, subkeyName, Length) Then Exit Do
        
        ' add to the result collection
        subkeyName = Left$(subkeyName, InStr(subkeyName, vbNullChar) - 1)
        EnumRegistryKeys.Add subkeyName, subkeyName
        ' prepare to query for next key
        Index = Index + 1
    Loop
   
    ' Close the key, if it was actually opened
    If handle Then RegCloseKey handle
        
End Function

