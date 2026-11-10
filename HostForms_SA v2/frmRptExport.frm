VERSION 5.00
Begin VB.Form frmRptExport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Export Form"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFilename 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1050
      TabIndex        =   7
      Top             =   900
      Width           =   6465
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   1890
      TabIndex        =   5
      Top             =   1410
      Width           =   1905
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7170
      TabIndex        =   4
      Top             =   480
      Width           =   315
   End
   Begin VB.TextBox txtLocation 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   6105
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
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
      Left            =   3960
      TabIndex        =   2
      Top             =   1410
      Width           =   1905
   End
   Begin VB.ComboBox cboFileType 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   6435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "File Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   270
      TabIndex        =   6
      Top             =   540
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   795
   End
End
Attribute VB_Name = "frmRptExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private yFilename As String
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
Private Const CSIDL_DESKTOP = &H0 '// The Desktop - virtual folder
Private Const CSIDL_PROGRAMS = 2 '// Program Files
Private Const CSIDL_CONTROLS = 3 '// Control Panel - virtual folder
Private Const CSIDL_PRINTERS = 4 '// Printers - virtual folder
Private Const CSIDL_DOCUMENTS = 5 '// My Documents
Private Const CSIDL_FAVORITES = 6 '// Favourites
Private Const CSIDL_STARTUP = 7 '// Startup Folder
Private Const CSIDL_RECENT = 8 '// Recent Documents
Private Const CSIDL_SENDTO = 9 '// Send To Folder
Private Const CSIDL_BITBUCKET = 10 '// Recycle Bin - virtual folder
Private Const CSIDL_STARTMENU = 11 '// Start Menu
Private Const CSIDL_DESKTOPFOLDER = 16 '// Desktop folder
Private Const CSIDL_DRIVES = 17 '// My Computer - virtual folder
Private Const CSIDL_NETWORK = 18 '// Network Neighbourhood - virtual folder
Private Const CSIDL_NETHOOD = 19 '// NetHood Folder
Private Const CSIDL_FONTS = 20 '// Fonts folder
Private Const CSIDL_SHELLNEW = 21 '// ShellNew folder
   
Private Sub fillcboExport()
Dim DefaultType As String
DefaultType = GetSetting("Lab-I", "ReportExport", "FileType", "Rich Text Format (RTF)")
With cboFileType
    .AddItem ""
    .AddItem "Comma Separated Values (CSV)"
    .AddItem "Excel (XLS)"
    .AddItem "Rich Text Format (RTF)"
    .AddItem "HTML"
    .AddItem "Word for Windows (DOC)"
    .AddItem "Text File"
    .ListIndex = 0
End With
findCboIndex cboFileType, DefaultType
End Sub

Private Sub cboFileType_Change()
    SaveSetting "Lab-I", "ReportExport", "FileType", cboFileType.Text
    ValidateValues
End Sub

Private Sub cboFileType_Click()
    cboFileType_Change
End Sub

Private Sub cmdBrowse_Click()
   txtLocation.Text = BrowseForFolder(Me.hWnd, "Browse")
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
    ReturnValueStr = cboFileType.Text & "|" & txtLocation.Text & "|" & txtFilename.Text
    Unload Me
End Sub

Private Sub Form_Load()
    fillcboExport
    txtFilename.Text = yFilename
    txtLocation.Text = GetSetting("Lab-I", "ReportExport", "FileLocation", fGetSpecialFolder(Me, 5))
End Sub


Public Function BrowseForFolder(ByVal lngHwnd As Long, ByVal strPrompt As String) As String

    On Error GoTo ehBrowseForFolder 'Trap for errors

    Dim intNull As Integer
    Dim lngIDList As Long, lngResult As Long
    Dim strPath As String
    Dim udtBI As BrowseInfo

    'Set API properties (housed in a UDT)
    With udtBI
        .lngHwnd = lngHwnd
        .lpszTitle = lstrcat(strPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Display the browse folder...
    lngIDList = SHBrowseForFolder(udtBI)

    If lngIDList <> 0 Then
        'Create string of nulls so it will fill in with the path
        strPath = String(MAX_PATH, 0)

        'Retrieves the path selected, places in the null
         'character filled string
        lngResult = SHGetPathFromIDList(lngIDList, strPath)

        'Frees memory
        Call CoTaskMemFree(lngIDList)

        'Find the first instance of a null character,
         'so we can get just the path
        intNull = InStr(strPath, vbNullChar)
        'Greater than 0 means the path exists...
        If intNull > 0 Then
            'Set the value
            strPath = Left(strPath, intNull - 1)
        End If
    End If

    'Return the path name
    BrowseForFolder = strPath
    Exit Function 'Abort

ehBrowseForFolder:

    'Return no value
    BrowseForFolder = Empty

End Function

Public Property Let wFilename(ByVal vNewValue As String)
    yFilename = vNewValue
End Property

Private Function findCboIndex(cboTemp As ComboBox, texttofind As String) As Integer
Dim ctr As Integer
With cboTemp
    For ctr = 0 To .ListCount - 1
        .ListIndex = ctr
        If .Text = texttofind Then
            findCboIndex = .ListIndex
            Exit Function
        End If
    Next ctr
    findCboIndex = .ListIndex = -1
End With
End Function

Private Sub ValidateValues()
If Len(Trim(cboFileType.Text)) = 0 Or Len(Trim(txtLocation.Text)) = 0 Or Len(Trim(txtFilename.Text)) = 0 Then
    cmdExport.Enabled = False
Else
    cmdExport.Enabled = True
End If
End Sub

Private Sub txtFilename_Change()
    ValidateValues
End Sub

Private Sub txtLocation_Change()
    SaveSetting "Lab-I", "ReportExport", "FileLocation", txtLocation.Text
    ValidateValues
End Sub
