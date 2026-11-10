Attribute VB_Name = "ModFunctions"
Option Explicit

'// API Declarations
Public Declare Function CreateFile Lib "Kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long

'// API Structures
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

'// API constants
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3
Public Const FILE_ATTRIBUTE_NORMAL = &H80

Public Declare Function GetVersionExA Lib "Kernel32" _
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


    
Private Type BrowseInfo
    lngHwnd        As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32.dll" _
    (ByVal hMem As Long)

Private Declare Function lstrcat Lib "Kernel32" _
   Alias "lstrcatA" (ByVal lpString1 As String, _
   ByVal lpString2 As String) As Long
   
Private Declare Function SHBrowseForFolder Lib "shell32" _
   (lpbi As BrowseInfo) As Long
   
Private Declare Function SHGetPathFromIDList Lib "shell32" _
   (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public returnValue As String
 

'// Return TRUE if the COM exists, FALSE if the COM does not exist
Public Function COMAvailable(COMNum As Integer) As Boolean
    Dim hCOM As Long
    Dim ret As Long
    Dim sec As SECURITY_ATTRIBUTES

    '// try to open the COM port
    hCOM = CreateFile("\.\COM" & COMNum & "", 0, FILE_SHARE_READ + _
        FILE_SHARE_WRITE, sec, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hCOM = -1 Then
        COMAvailable = False
    Else
        COMAvailable = True
        '// close the COM port
        ret = CloseHandle(hCOM)
    End If
End Function

Public Function getVersion() As String
Dim OSInfo As OSVERSIONINFO
Dim retvalue As Integer
  
OSInfo.dwOSVersionInfoSize = 148
OSInfo.szCSDVersion = Space$(128)
retvalue = GetVersionExA(OSInfo)
  
With OSInfo
     Select Case .dwPlatformId
         Case VER_PLATFORM_WIN32s           ' Win32s on Windows 3.1
              getVersion = "WIN32" '"Windows 3.1"
               
         Case VER_PLATFORM_WIN32_WINDOWS    ' Windows 95, Windows 98,
              Select Case .dwMinorVersion   ' or Windows Me
                  Case 0
                      getVersion = "WIN32" '"Windows 95"
                  Case 10
                      If (OSInfo.dwBuildNumber And &HFFFF&) = 2222 Then
                          getVersion = "WIN32" '"Windows 98SE"
                      Else
                          getVersion = "WIN32" '"Windows 98"
                      End If
                  Case 90
                      getVersion = "WIN32" '"Windows Me"
              End Select
     
         Case VER_PLATFORM_WIN32_NT         ' Windows NT, Windows 2000, Windows XP,
              Select Case .dwMajorVersion   ' or Windows Server 2003 family.
                  Case 3
                      getVersion = "NT" '"Windows NT 3.51"
                  Case 4
                      getVersion = "NT" '"Windows NT 4.0"
                  Case 5, 6
                      Select Case .dwMinorVersion
                          Case 0
                              getVersion = "NT" '"Windows 2000"
                          Case 1
                              getVersion = "NT" '"Windows XP"
                          Case 2
                              getVersion = "NT" '"Windows Server 2003"
                      End Select
                  Case 10
                      Select Case .dwMinorVersion
                          Case 0
                              getVersion = "NT" '"Windows 2000"
                          Case Else
                              getVersion = "NT"
                      End Select
              End Select
         Case Else
              getVersion = "Failed"
               
     End Select
             
End With
End Function

Public Sub setFormColors(myForm As Form)
Dim ctr As Integer
On Error Resume Next
    myForm.fraTitle.BackColor = GetSetting("Lab-I", "FormColors", "TitleBar", &HA66B2B)
    myForm.lblTitle.ForeColor = GetSetting("Lab-I", "FormColors", "TitleText", &HFFFFFF)
    For ctr = 0 To myForm.fraBody.UBound
        myForm.fraBody(ctr).BackColor = GetSetting("Lab-I", "FormColors", "BodyFrame", &H6B4B01)
        myForm.fraBody(ctr).ForeColor = GetSetting("Lab-I", "FormColors", "BodyText", &HFFFFFF)
    Next ctr
    myForm.lblStatus.ForeColor = GetSetting("Lab-I", "FormColors", "BodyText", &HFFFFFF)
    For ctr = 0 To myForm.lblText.UBound
        myForm.lblText(ctr).ForeColor = GetSetting("Lab-I", "FormColors", "BodyText", &HFFFFFF)
    Next ctr
    For ctr = 0 To myForm.chkValues.UBound
        myForm.chkValues(ctr).ForeColor = GetSetting("Lab-I", "FormColors", "BodyText", &HFFFFFF)
        myForm.chkValues(ctr).BackColor = GetSetting("Lab-I", "FormColors", "BodyFrame", &H6B4B01)
    Next ctr
    myForm.chkEnabled.ForeColor = GetSetting("Lab-I", "FormColors", "BodyText", &HFFFFFF)
    myForm.chkEnabled.BackColor = GetSetting("Lab-I", "FormColors", "BodyFrame", &H6B4B01)
End Sub

