VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmOBISetiings 
   Caption         =   "Database"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   10965
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTimeOut 
      Height          =   315
      Left            =   3360
      TabIndex        =   36
      Top             =   6870
      Width           =   855
   End
   Begin VB.CheckBox chkShowErrorMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Show Error Message"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   6930
      Width           =   2115
   End
   Begin VB.CheckBox chkExternal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Use External Database"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   3990
      Width           =   3345
   End
   Begin VB.Frame fraRestApi 
      Caption         =   "API Values"
      Height          =   2715
      Left            =   120
      TabIndex        =   21
      Top             =   1170
      Visible         =   0   'False
      Width           =   3585
      Begin VB.Timer tmrTimeOut 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2670
         Top             =   2190
      End
      Begin VB.TextBox txtRestTable 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   630
         TabIndex        =   33
         Top             =   1140
         Width           =   2865
      End
      Begin VB.TextBox txtTestValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   31
         Top             =   1740
         Width           =   1695
      End
      Begin VB.TextBox txtFieldName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   30
         Top             =   1740
         Width           =   1665
      End
      Begin VB.ComboBox cboMethod 
         Height          =   315
         ItemData        =   "frmOBISetiings.frx":0000
         Left            =   1260
         List            =   "frmOBISetiings.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox cboContentType 
         Height          =   315
         ItemData        =   "frmOBISetiings.frx":0019
         Left            =   1260
         List            =   "frmOBISetiings.frx":0023
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   630
         Width           =   2295
      End
      Begin MSWinsockLib.Winsock wsTest 
         Left            =   3120
         Top             =   2190
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton cmdTestApi 
         Caption         =   "Test API"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   90
         TabIndex        =   23
         Top             =   2100
         Width           =   3405
      End
      Begin VB.Label Label10 
         Caption         =   "Table"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   32
         Top             =   1170
         Width           =   945
      End
      Begin VB.Label Label9 
         Caption         =   "Value"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1830
         TabIndex        =   29
         Top             =   1500
         Width           =   705
      End
      Begin VB.Label Label8 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   28
         Top             =   1500
         Width           =   645
      End
      Begin VB.Line Line1 
         X1              =   60
         X2              =   3540
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Label Label7 
         Caption         =   "Content Type"
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
         Left            =   90
         TabIndex        =   27
         Top             =   660
         Width           =   1125
      End
      Begin VB.Label Label6 
         Caption         =   "Method"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   540
         TabIndex        =   26
         Top             =   300
         Width           =   945
      End
   End
   Begin VB.Frame fraServerName 
      Caption         =   "SERVER NAME"
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   90
      TabIndex        =   18
      Top             =   420
      Width           =   3615
      Begin VB.TextBox txtServerName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   90
         TabIndex        =   19
         Top             =   270
         Width           =   3465
      End
   End
   Begin VB.ComboBox cboServerType 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   90
      Width           =   2385
   End
   Begin VB.Frame fraServer 
      Caption         =   "CONNECTION SETTINGS"
      ForeColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   90
      TabIndex        =   7
      Top             =   1140
      Width           =   3615
      Begin VB.TextBox txtTable 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         TabIndex        =   12
         Top             =   1500
         Width           =   2385
      End
      Begin VB.TextBox txtDatabase 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         TabIndex        =   11
         Top             =   1110
         Width           =   2385
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1110
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   720
         Width           =   2385
      End
      Begin VB.TextBox txtUsername 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         TabIndex        =   9
         Top             =   330
         Width           =   2385
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   375
         Left            =   1110
         TabIndex        =   8
         Top             =   1950
         Width           =   2385
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Table:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   450
         TabIndex        =   16
         Top             =   1530
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   270
         TabIndex        =   15
         Top             =   390
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   300
         TabIndex        =   14
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Database:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   360
         TabIndex        =   13
         Top             =   1140
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4155
      Left            =   3750
      TabIndex        =   4
      Top             =   90
      Width           =   7185
      Begin VB.ComboBox cboList 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3300
         Visible         =   0   'False
         Width           =   2595
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   3675
         Left            =   90
         TabIndex        =   6
         Top             =   180
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   6482
         _Version        =   393216
         BackColor       =   16777215
         FixedCols       =   0
         RowHeightMin    =   360
         BackColorFixed  =   10906411
         ForeColorFixed  =   16777215
         BackColorBkg    =   12632256
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   345
      Left            =   8610
      TabIndex        =   3
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   345
      Left            =   9780
      TabIndex        =   2
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Criteria"
      ForeColor       =   &H00FFFFFF&
      Height          =   2445
      Left            =   90
      TabIndex        =   0
      Top             =   4290
      Width           =   10815
      Begin RichTextLib.RichTextBox rtbCriteria 
         Height          =   2055
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   3625
         _Version        =   393217
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmOBISetiings.frx":005C
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Timeout(sec):"
      Height          =   255
      Left            =   2340
      TabIndex        =   35
      Top             =   6900
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SERVER TYPE:"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmOBISetiings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private myCon As New ADODB.Connection
Private blnConnected As Boolean
Private TimeOutCtr As Integer

Private Sub cboList_Click()
With msfgResult
    .TextMatrix(.Row, .Col) = cboList.Text
End With
End Sub

Private Sub cboServerType_Change()
    cboServerType_Click
End Sub

Private Sub cboServerType_Click()
    If cboServerType.Text = "REST API" Then
        fraRestApi.Visible = True
        fraServer.Visible = False
        fraServerName.Caption = "URI"
        Frame3.Caption = "Result"
    Else
        fraRestApi.Visible = False
        fraServer.Visible = True
        fraServerName.Caption = "SERVER NAME"
        Frame3.Caption = "Criteria"
    End If
    rtbCriteria.Text = ""
    txtServerName.Text = ""
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdConnect_Click()
On Error GoTo errFound
    If cboServerType.Text = "Informix-CLI 2.5" Then
        myCon.ConnectionString = "Driver={Informix-CLI 2.5 (32 Bit)};Server=" & txtServerName & ";Database=" & txtDatabase & ";Uid=" & txtUsername & ";Pwd=" & txtPassword & ";"
    ElseIf cboServerType.Text = "MySQL" Then
        myCon.ConnectionString = "Driver={MySQL ODBC 3.51 Driver};Port=3306;Server=" & txtServerName & ";Database=" & txtDatabase & ";User=" & txtUsername & ";Password=" & txtPassword & ";Option=3;"
    ElseIf cboServerType.Text = "MS SQL Server 2000" Then
        myCon.ConnectionString = "Provider=sqloledb;Data Source=" & txtServerName & ";Initial Catalog=" & txtDatabase & ";Integrated Security=SSPI;User Id=" & txtUsername & ";Password=" & txtPassword
    End If
    myCon.ConnectionTimeout = 30
    myCon.CursorLocation = adUseClient
    myCon.Open
errFound:
    If Err.Number <> 0 Then
        Select Case Err.Number
            Case -2147467259
                MsgBox "Unable to connect to server" & vbCrLf & vbCrLf & "Please check database settings."
            Case Else
                MsgBox Err.Description & vbCrLf & "Please Contact your Systems Administrator."
        End Select
    Else
        MsgBox "Connection OK."
        FillCboList
    End If
If myCon.State = 1 Then myCon.Close
End Sub

Private Sub cmdSave_Click()
Dim ctr As Integer
    SaveSetting "LAB-I", "OBISettings", "ServerType", cboServerType.Text
    SaveSetting "LAB-I", "OBISettings", "ServerTypeIndex", cboServerType.ListIndex
    SaveSetting "LAB-I", "OBISettings", "ServerName", txtServerName.Text
    SaveSetting "LAB-I", "OBISettings", "UserName", txtUsername.Text
    SaveSetting "LAB-I", "OBISettings", "Password", txtPassword.Text
    SaveSetting "LAB-I", "OBISettings", "Database", txtDatabase.Text
    SaveSetting "LAB-I", "OBISettings", "Table", txtTable.Text
    SaveSetting "LAB-I", "OBISettings", "Criteria", rtbCriteria.Text
    SaveSetting "LAB-I", "OBISettings", "UseExDatabase", chkExternal.Value
    SaveSetting "LAB-I", "OBISettings", "RestApiMethod", cboMethod.Text
    SaveSetting "LAB-I", "OBISettings", "RestApiContentType", cboContentType
    SaveSetting "LAB-I", "OBISettings", "RestApiTestField", txtFieldName.Text
    SaveSetting "LAB-I", "OBISettings", "RestApiTestValue", txtTestValue.Text
    SaveSetting "LAB-I", "OBISettings", "RestApiTestTable", txtRestTable.Text
    SaveSetting "LAB-I", "OBISettings", "ShowErrorMessage", chkShowErrorMsg.Value
    SaveSetting "LAB-I", "OBISettings", "QueryTimeOut", txtTimeOut.Text
    With msfgResult
        For ctr = 1 To .Rows - 1
            SaveSetting "LAB-I", "OBISettings", "FLD_" & .TextMatrix(ctr, 0), IIf(.TextMatrix(ctr, 2) = "", "''", .TextMatrix(ctr, 2))
        Next ctr
    End With
    MsgBox "Settings successfully saved.", vbInformation + vbOKOnly
End Sub

Private Sub cmdTestApi_Click()
Me.MousePointer = 11
    cmdTestApi.Enabled = False
    GenerateTestPacket
    cmdTestApi.Enabled = True
Me.MousePointer = 0
End Sub

Private Sub Form_Load()
Dim ctr  As Integer
With cboServerType
    .AddItem "MS SQL Server 2000"
    .AddItem "MySQL"
    .AddItem "REST API"
End With
'cboServerType.Text = GetSetting("LAB-I", "OBISettings", "ServerType", "")
cboServerType.ListIndex = GetSetting("LAB-I", "OBISettings", "ServerTypeIndex", -1)
txtServerName.Text = GetSetting("LAB-I", "OBISettings", "ServerName", "")
txtUsername.Text = GetSetting("LAB-I", "OBISettings", "UserName", "")
txtPassword.Text = GetSetting("LAB-I", "OBISettings", "Password", "")
txtDatabase.Text = GetSetting("LAB-I", "OBISettings", "Database", "")
txtTable.Text = GetSetting("LAB-I", "OBISettings", "Table", "")
rtbCriteria.Text = GetSetting("LAB-I", "OBISettings", "Criteria", "")

cboMethod.ListIndex = findCboIndex(cboMethod, GetSetting("LAB-I", "OBISettings", "RestApiMethod", ""))
cboContentType.ListIndex = findCboIndex(cboContentType, GetSetting("LAB-I", "OBISettings", "RestApiContentType", ""))
txtFieldName.Text = GetSetting("LAB-I", "OBISettings", "RestApiTestField", "")
txtTestValue.Text = GetSetting("LAB-I", "OBISettings", "RestApiTestValue", "")
txtRestTable.Text = GetSetting("LAB-I", "OBISettings", "RestApiTestTable", "")
setmsfgResult
With msfgResult
    For ctr = 1 To .Rows - 1
        If GetSetting("LAB-I", "OBISettings", "FLD_" & ctr, "") <> "''" Then
            .TextMatrix(ctr, 2) = GetSetting("LAB-I", "OBISettings", "FLD_" & ctr, "")
        End If
    Next ctr
End With
chkExternal.Value = Val(GetSetting("LAB-I", "OBISettings", "UseExDatabase", "0"))
chkShowErrorMsg.Value = Val(GetSetting("LAB-I", "OBISettings", "ShowErrorMessage", "0"))
txtTimeOut.Text = GetSetting("LAB-I", "OBISettings", "QueryTimeOut", "5")
End Sub

Private Sub FillCboList()
Dim ctr As Integer
Dim rstTemp As ADODB.Recordset
Dim arrTemp() As String
Dim arrTemp2() As String
Dim tempstr As String
If cboServerType.Text = "MS SQL Server 2000" Then
    Set rstTemp = myCon.Execute("Select top 1 * from " & txtTable.Text)
ElseIf cboServerType.Text = "MySQL" Then
    Set rstTemp = myCon.Execute("Select * from " & txtTable.Text & " limit 1")
ElseIf cboServerType.Text = "REST API" Then
    tempstr = Replace(rtbCriteria.Text, """, """, """,""")
    tempstr = Replace(rtbCriteria.Text, """, " & vbCrLf & """", """,""")
    tempstr = Replace(rtbCriteria.Text, """: """, """:""")
    If InStr(1, tempstr, "{" & vbCrLf) > 0 Then
        tempstr = Trim(Mid(tempstr, InStr(1, tempstr, "{" & vbCrLf) + 3))
    End If
    If InStr(1, tempstr, vbCrLf & "}") > 0 Then
        tempstr = Trim(Mid(tempstr, 1, InStr(1, tempstr, vbCrLf & "}")))
    End If
    tempstr = Replace(tempstr, """,""", "|")
    tempstr = Replace(tempstr, """:""", "^")
    If Left(tempstr, 2) = "{""" And Right(tempstr, 2) = """}" Then
        tempstr = Mid(tempstr, 3)
        tempstr = Mid(tempstr, 1, Len(tempstr) - 2)
    End If
End If
If cboServerType.Text = "REST API" Then
    arrTemp = Split(tempstr, "|")
    With cboList
        .Clear
        .AddItem ""
        For ctr = 0 To UBound(arrTemp)
            arrTemp2 = Split(arrTemp(ctr), "^")
            .AddItem arrTemp2(0)
        Next ctr
    End With
Else
    With cboList
        .Clear
        .AddItem ""
        For ctr = 0 To rstTemp.Fields.Count - 1
            .AddItem rstTemp.Fields(ctr).Name
        Next ctr
    End With
    Set rstTemp = Nothing
End If
End Sub

Private Sub setmsfgResult()
Dim ctr As Integer
With msfgResult
    .Cols = 3
    .Rows = 20
    .ColWidth(0) = .Width * 0.25 - (30 * 3)
    .ColWidth(1) = .Width * 0.3 - (30 * 3)
    .ColWidth(2) = .Width * 0.45 - (50 * 3)
    .TextMatrix(0, 0) = "Lable ID"
    .TextMatrix(0, 1) = "Lable"
    .TextMatrix(0, 2) = "Assigned Field"
    .TextMatrix(1, 1) = GetSetting("HostForms_SA", "LabelSettings", "PatientID", "Patient ID")
    .TextMatrix(2, 1) = GetSetting("HostForms_SA", "LabelSettings", "SpecimenID", "Specimen ID")
    .TextMatrix(3, 1) = "Last Name"
    .TextMatrix(4, 1) = "First Name"
    .TextMatrix(5, 1) = "Middle Name"
    .TextMatrix(6, 1) = GetSetting("HostForms_SA", "LabelSettings", "Age", "Age")
    .TextMatrix(7, 1) = GetSetting("HostForms_SA", "LabelSettings", "Sex", "Sex")
    .TextMatrix(8, 1) = GetSetting("HostForms_SA", "LabelSettings", "Ward", "Ward")
    .TextMatrix(9, 1) = GetSetting("HostForms_SA", "LabelSettings", "Room", "Room")
    
    .TextMatrix(10, 1) = GetSetting("HostForms_SA", "LabelSettings", "Resultdate", "Result Date")
    .TextMatrix(11, 1) = .TextMatrix(10, 1) & " - Time"
    .TextMatrix(12, 1) = GetSetting("HostForms_SA", "LabelSettings", "RequestDate", "Request Date")
    .TextMatrix(13, 1) = .TextMatrix(12, 1) & " - Time"
    
    .TextMatrix(14, 1) = GetSetting("HostForms_SA", "LabelSettings", "Physician", "Physician")
    .TextMatrix(15, 1) = GetSetting("HostForms_SA", "LabelSettings", "Medtech", "Med. Tech.")
    .TextMatrix(16, 1) = GetSetting("HostForms_SA", "LabelSettings", "Remarks", "Remarks")
    .TextMatrix(17, 1) = GetSetting("HostForms_SA", "LabelSettings", "BirthDate", "Birth Date")
    .TextMatrix(18, 1) = GetSetting("HostForms_SA", "LabelSettings", "CivilStatus", "Civil Status")
    .TextMatrix(19, 1) = GetSetting("HostForms_SA", "LabelSettings", "Specimen", "Specimen")
    For ctr = 1 To 19
        .TextMatrix(ctr, 0) = ctr
    Next ctr
End With
End Sub

Private Sub msfgResult_Click()
With msfgResult
    .Col = 2
    cboList.Visible = True
    cboList.Width = .CellWidth
    cboList.Top = .CellTop + .Top
    cboList.Left = .CellLeft + .Left
    If cboList.ListCount > 1 Then findcboType .TextMatrix(.Row, .Col)
End With
End Sub

Private Sub findcboType(myvalue As String)
Dim ctr As Integer
Dim mFound As Boolean
Dim tIndex As Integer
mFound = False
With cboList
    For ctr = 0 To .ListCount - 1
        .ListIndex = ctr
        If .Text = myvalue Then
            mFound = True
            tIndex = .ListIndex
            Exit For
        End If
    Next ctr
If mFound Then .ListIndex = tIndex
End With
End Sub

Private Sub msfgResult_Scroll()
    cboList.Visible = False
End Sub

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
End With
End Function

Private Function GenerateTestPacket() As String
Dim tempstr As String
Dim mPort As String
Dim mAddress As String
Dim mPath As String
Dim tempPacket As String
Dim arrTemp() As String
Dim mData As String
If Format(Left(Trim(txtServerName.Text), 7), ">") = "HTTP://" Then
    tempstr = Mid(Trim(txtServerName.Text), 8)
    mPath = Mid(tempstr, InStr(1, tempstr, "/"))
    tempstr = Mid(tempstr, 1, InStr(1, tempstr, "/") - 1)
Else
    tempstr = Mid(txtServerName.Text, 1, InStr(1, txtServerName.Text, "/") - 1)
    mPath = Mid(txtServerName.Text, InStr(1, txtServerName.Text, "/"))
End If
arrTemp = Split(tempstr, ":")
If InStr(1, tempstr, ":") > 0 Then
    mPort = arrTemp(1)
    mAddress = arrTemp(0)
Else
    mPort = "80"
    mAddress = arrTemp(0)
End If
tempPacket = ""
tempPacket = tempPacket & "Host: " & mAddress & vbCrLf
tempPacket = tempPacket & "Accept: */*" & vbCrLf
'tempPacket = tempPacket & "Expect: 100-Continue" & vbCrLf
tempPacket = tempPacket & "Content-Type: " & cboContentType.Text & vbCrLf

Select Case Format(cboContentType.Text, ">")
    Case "APPLICATION/JSON"
        mData = "{""" & txtFieldName.Text & """:""" & txtTestValue.Text & """}"
        If cboMethod.Text = "GET" Then
            If Len(Trim(txtFieldName.Text)) = 0 Then
                mData = ""
            Else
                mData = "/" & txtFieldName.Text & "/" & txtTestValue.Text
            End If
            tempPacket = cboMethod.Text & " " & mPath & mData & " HTTP/1.1" & vbCrLf & tempPacket & vbCrLf & vbCrLf
        Else
            tempPacket = cboMethod.Text & " " & mPath & " HTTP/1.1" & vbCrLf & tempPacket
            tempPacket = tempPacket & "Content-Length: " & Len(mData) & vbCrLf & vbCrLf & mData
        End If
    Case "APPLICATION/X-WWW-FORM-URLENCODED"
        mData = txtFieldName.Text & "=" & txtTestValue.Text
        tempPacket = cboMethod.Text & " " & mPath & "?" & mData & " HTTP/1.1" & vbCrLf & tempPacket
        If cboMethod.Text = "POST" Then
            tempPacket = tempPacket & "Content-Length: " & Len(mData) & vbCrLf & vbCrLf & mData
        End If
    Case Else
        tempPacket = cboMethod.Text & " " & mPath & " HTTP/1.1" & vbCrLf & tempPacket
        mData = ""
End Select
wsTest.Close
blnConnected = False
wsTest.Protocol = sckTCPProtocol
wsTest.RemoteHost = mAddress
wsTest.RemotePort = mPort
wsTest.Connect
TimeOutCtr = 0
tmrTimeout.Enabled = True
While Not blnConnected And TimeOutCtr <= Val(txtTimeOut.Text)
    DoEvents
Wend
If blnConnected Then
    wsTest.SendData tempPacket
Else
    TimeOutCtr = 0
    tmrTimeout.Enabled = False
    rtbCriteria.Text = "{""ERROR"":""Server cannot be found...""}"
    If chkShowErrorMsg.Value = 1 And GetError(rtbCriteria.Text, 1) = "1" Then
        MsgBox GetError(rtbCriteria.Text, 3), vbOKOnly + vbInformation, "PID Search"
    End If
End If
End Function

Private Sub rtbCriteria_Change()
If cboServerType.Text = "REST API" Then
    FillCboList
End If
End Sub

Private Sub tmrTimeOut_Timer()
TimeOutCtr = TimeOutCtr + 1
If TimeOutCtr > Val(txtTimeOut.Text) Then tmrTimeout.Enabled = False
End Sub

Private Sub wsTest_Close()
    blnConnected = False
    wsTest.Close
End Sub

Private Sub wsTest_Connect()
    blnConnected = True
End Sub

Private Sub wsTest_DataArrival(ByVal bytesTotal As Long)
Dim strPageReturn As String
wsTest.GetData strPageReturn
rtbCriteria.Text = ProcessReceived(strPageReturn)
End Sub

Private Sub wsTest_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
wsTest.Close
End Sub

Private Function ProcessReceived(mess As String) As String
Dim mData As String
If GetStatusCode(mess) = "200" Then
    If InStr(1, mess, vbCrLf & vbCrLf) > 0 Then
        mData = Mid(mess, InStr(1, mess, vbCrLf & vbCrLf) + 4)
    Else
        mData = mess
    End If
    mData = Replace(mData, """, """, """,""")
    If InStr(1, mData, "{" & vbCrLf) > 0 Then
        mData = Mid(mData, InStr(1, mData, "{" & vbCrLf) + 3)
    End If
    If InStr(1, mData, vbCrLf & "}") > 0 Then
        mData = Mid(mData, 1, InStr(1, mData, vbCrLf & "}"))
    End If
    
    mData = Replace(mData, """: [", """:[")
    If InStr(1, mData, """" & txtRestTable.Text & """:[", vbTextCompare) > 0 Then
        mData = Mid(mData, InStr(1, mData, """" & txtRestTable.Text & """:[", vbTextCompare) + Len("""" & txtRestTable.Text & """:["))
        mData = Mid(mData, 1, InStr(1, mData, "]"))
        mData = Mid(mData, 1, InStr(1, mData, "}"))
    End If
    mData = Replace(mData, """:NULL,", """:"""",", , , vbTextCompare)
    'Trim Carriage Return
    Do While Left(mData, 2) = vbCrLf
        mData = Mid(mData, 3)
    Loop
    Do While Right(mData, 2) = vbCrLf
        mData = Mid(mData, 1, Len(mData) - 2)
    Loop
    '--------------
    mData = Trim(mData)
    If mData = "[]" Then
        mData = "{""ERROR"":""NO RECORD FOUND""}"
    End If
Else
    mData = "{""ERROR"":""" & GetStatusCode(mess) & ": " & GetHTTPStatusCodeDesc(Val(GetStatusCode(mess)), False) & """}"
End If
If chkShowErrorMsg.Value = 1 And GetError(mData, 1) = "1" Then
    MsgBox GetError(mData, 3), vbOKOnly + vbInformation, "PID Search"
End If
ProcessReceived = mData
End Function

Private Function GetStatusCode(mess As String) As String
Dim arrTemp() As String
Dim tempstr As String
tempstr = ""
If Len(mess) > 0 And InStr(1, mess, " ") Then
    tempstr = Mid(mess, 1, InStr(1, mess, vbCrLf))
    tempstr = Replace(tempstr, "   ", " ")
    tempstr = Replace(tempstr, "  ", " ")
    arrTemp = Split(tempstr, " ")
    If UBound(arrTemp) >= 1 Then
        tempstr = arrTemp(1)
    End If
End If
GetStatusCode = tempstr
End Function

Private Function GetHTTPStatusCodeDesc(ByVal pintHTTPStatusCode As Integer, pbolVerbose As Boolean) As String
    '*******************************************************************************
    '
    ' Function       : GetHTTPStatusCodeDesc
    '
    ' Return Type    : String
    '
    ' Copyright      : 2004 - Salty Brine Software (Art Araya)
    '
    ' Description    : Provides either a brief or verbose description for an
    '                  HTTP Status Code.
    '
    ' Parameters     : 1) pintHTTPStatusCode (Integer) - the HTTP Status Code
    '                     that a desciption Is needed For
    '                : 2) pbolVerbose (Boolean) - Used to specify whether a brief or
    '                     verbose description of the status code Is desired.
    '
    ' Range Description
    ' 100-199 Informational
    ' 200-299 Success: indicates that a client's request arrived and was accepted.
    ' 300-399 Client request redirection, further action required.
    ' 400-499 Client error
    ' 500-599 Server error
    '
    '*******************************************************************************

    Dim strHTTPStatusCodeDesc   As String

    If pbolVerbose Then

        Select Case pintHTTPStatusCode
            Case 100
                strHTTPStatusCodeDesc = "Continue - The client can continue its request."
            Case 101
                strHTTPStatusCodeDesc = "Switching Protocols - The server has complied with the client's request to switch protocols."
            Case 200
                strHTTPStatusCodeDesc = "OK - The request was successful."
            Case 201
                strHTTPStatusCodeDesc = "Created - The request was successful and a new resource was created as a result."
            Case 202
                strHTTPStatusCodeDesc = "Accepted - The request was accepted for processing."
            Case 203
                strHTTPStatusCodeDesc = "Non-Authoritative Information - The meta-information in the entity-header is from a local or third-party copy."
            Case 204
                strHTTPStatusCodeDesc = "No Content - The server serviced the request but there is no information to return."
            Case 205
                strHTTPStatusCodeDesc = "Reset Content - The request was successful and the user agent should clear the input form for further input."
            Case 206
                strHTTPStatusCodeDesc = "Partial Content - The server serviced a (byte) range request for the resource."
            Case 300
                strHTTPStatusCodeDesc = "Multiple Choices - The requested resource could be one of multiple resources."
            Case 301
                strHTTPStatusCodeDesc = "Moved Permanently - The requested resource has permanently moved to a new location."
            Case 302
                strHTTPStatusCodeDesc = "Moved Temporarily - The requested resource has temporarily moved to a new location."
            Case 303
                strHTTPStatusCodeDesc = "See Other - The response to the request is under a different URI and can be retrieved with a GET request."
            Case 304
                strHTTPStatusCodeDesc = "Not Modified - The requested resource has not been modified since it was last requested."
            Case 305
                strHTTPStatusCodeDesc = "Use Proxy - The requested resource must be accessed through a proxy server."
            Case 400
                strHTTPStatusCodeDesc = "Bad Request - The server cannot read the request because its syntax is incorrect."
            Case 401
                strHTTPStatusCodeDesc = "Unauthorized - The server must authenticate the user before servicing the request."
            Case 402
                strHTTPStatusCodeDesc = "Payment Required"
            Case 403
                strHTTPStatusCodeDesc = "Forbidden - The server refused to service the request."
            Case 404
                strHTTPStatusCodeDesc = "Not Found - The server cannot find the requested resource."
            Case 405
                strHTTPStatusCodeDesc = "Method Not Allowed - The method specified in the Request-Line is not permitted for the requested resource."
            Case 406
                strHTTPStatusCodeDesc = "Not Acceptable - The requested resource can only generate response entities that have unacceptable content characteristics according to the accept headers sent in the request."
            Case 407
                strHTTPStatusCodeDesc = "Proxy Authentication Required - The proxy server must authenticate the user before servicing the request."
            Case 408
                strHTTPStatusCodeDesc = "Request Timeout - The client did not make its request within the amount of time the server will wait for requests."
            Case 409
                strHTTPStatusCodeDesc = "Conflict - The server could not service the request due to a current conflict with the requested resource."
            Case 410
                strHTTPStatusCodeDesc = "Gone - The requested resource is no longer available on the server."
            Case 411
                strHTTPStatusCodeDesc = "Length Required - The server will not service the request without a Content-Length specified in the request."
            Case 412
                strHTTPStatusCodeDesc = "Precondition Failed - A precondition specified in one or more of the request-header fields failed."
            Case 413
                strHTTPStatusCodeDesc = "Request Entity Too Large - The server will not service the request because the requested resource is too large."
            Case 414
                strHTTPStatusCodeDesc = "Request-URI Too Long - The server will not service the request because the requested URL is too long."
            Case 415
                strHTTPStatusCodeDesc = "Unsupported Media Type - The server will not service the request because the format of the request is not supported by the requested resource for the requested method."
            Case 500
                strHTTPStatusCodeDesc = "Internal Server Error - The server could not service the request because of an unexpected internal error."
            Case 501
                strHTTPStatusCodeDesc = "Not Implemented - The sever cannot service the request because it does not support the request method."
            Case 502
                strHTTPStatusCodeDesc = "Bad Gateway - The proxy server received an invalid response from the content server or another proxy server in a proxy chain."
            Case 503
                strHTTPStatusCodeDesc = "Service Unavailable - The server could not service the request because it was temporarily overloaded or undergoing maintenance."
            Case 504
                strHTTPStatusCodeDesc = "Gateway Timeout - The proxy server did not receive a response from a chained proxy server or the origin content server within an acceptable amount of time."
            Case 505
                strHTTPStatusCodeDesc = "HTTP Version Not Supported - The server does not support the HTTP version specified in the request."
            Case Else
                strHTTPStatusCodeDesc = "Unrecognized HTTP Status Code"
        End Select

    Else

        Select Case pintHTTPStatusCode
            Case 100
                strHTTPStatusCodeDesc = "Continue"
            Case 101
                strHTTPStatusCodeDesc = "Switching Protocols"
            Case 200
                strHTTPStatusCodeDesc = "OK"
            Case 201
                strHTTPStatusCodeDesc = "Created"
            Case 202
                strHTTPStatusCodeDesc = "Accepted"
            Case 203
                strHTTPStatusCodeDesc = "Non-Authoritative Information"
            Case 204
                strHTTPStatusCodeDesc = "No Content"
            Case 205
                strHTTPStatusCodeDesc = "Reset Content"
            Case 206
                strHTTPStatusCodeDesc = "Partial Content"
            Case 300
                strHTTPStatusCodeDesc = "Multiple Choices"
            Case 301
                strHTTPStatusCodeDesc = "Moved Permanently"
            Case 302
                strHTTPStatusCodeDesc = "Moved Temporarily"
            Case 303
                strHTTPStatusCodeDesc = "See Other"
            Case 304
                strHTTPStatusCodeDesc = "Not Modified"
            Case 305
                strHTTPStatusCodeDesc = "Use Proxy"
            Case 400
                strHTTPStatusCodeDesc = "Bad Request"
            Case 401
                strHTTPStatusCodeDesc = "Unauthorized"
            Case 402
                strHTTPStatusCodeDesc = "Payment Required"
            Case 403
                strHTTPStatusCodeDesc = "Forbidden"
            Case 404
                strHTTPStatusCodeDesc = "Not Found"
            Case 405
                strHTTPStatusCodeDesc = "Method Not Allowed"
            Case 406
                strHTTPStatusCodeDesc = "Not Acceptable"
            Case 407
                strHTTPStatusCodeDesc = "Proxy Authentication Required"
            Case 408
                strHTTPStatusCodeDesc = "Request Timeout"
            Case 409
                strHTTPStatusCodeDesc = "Conflict"
            Case 410
                strHTTPStatusCodeDesc = "Gone"
            Case 411
                strHTTPStatusCodeDesc = "Length Required"
            Case 412
                strHTTPStatusCodeDesc = "Precondition Failed"
            Case 413
                strHTTPStatusCodeDesc = "Request Entity Too Large"
            Case 414
                strHTTPStatusCodeDesc = "Request-URI Too Long"
            Case 415
                strHTTPStatusCodeDesc = "Unsupported Media Type"
            Case 500
                strHTTPStatusCodeDesc = "Internal Server Error"
            Case 501
                strHTTPStatusCodeDesc = "Not Implemented"
            Case 502
                strHTTPStatusCodeDesc = "Bad Gateway"
            Case 503
                strHTTPStatusCodeDesc = "Service Unavailable"
            Case 504
                strHTTPStatusCodeDesc = "Gateway Timeout"
            Case 505
                strHTTPStatusCodeDesc = "HTTP Version Not Supported"
            Case Else
                strHTTPStatusCodeDesc = "Unrecognized HTTP Status Code"
        End Select
    End If  'If pbolVerbose

    GetHTTPStatusCodeDesc = strHTTPStatusCodeDesc

End Function

Private Function GetError(mess As String, mTrans As Integer) As String
Dim arrTemp() As String
Dim tempstr As String
Dim hasError As String
tempstr = Replace(mess, "{""", "")
tempstr = Replace(tempstr, """}", "")
tempstr = Replace(tempstr, """:""", "|")
arrTemp = Split(tempstr, "|")
hasError = "0"
If UBound(arrTemp) >= 1 Then
    If arrTemp(0) = "ERROR" Then
        hasError = "1"
    End If
End If
If hasError = "1" Then
    Select Case mTrans
        Case 1: GetError = hasError
        Case 2: GetError = arrTemp(0)
        Case 3: GetError = arrTemp(1)
    End Select
Else
    Select Case mTrans
        Case 1: GetError = hasError
        Case 2: GetError = ""
        Case 3: GetError = ""
    End Select
End If
End Function
