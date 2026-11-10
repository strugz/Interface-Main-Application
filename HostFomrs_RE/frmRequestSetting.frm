VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRequestSetting 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Request Settings"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   9135
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   -75
      Width           =   9960
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   405
         Left            =   7770
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   8625
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   8820
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   8625
         Width           =   975
      End
      Begin TabDlg.SSTab SSTab 
         Height          =   8280
         Left            =   210
         TabIndex        =   3
         Top             =   240
         Width           =   9585
         _ExtentX        =   16907
         _ExtentY        =   14605
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "Request Fields"
         TabPicture(0)   =   "frmRequestSetting.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraDemo"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Others"
         TabPicture(1)   =   "frmRequestSetting.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraGlobalOption"
         Tab(1).ControlCount=   1
         Begin VB.Frame fraGlobalOption 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   7800
            Left            =   -74820
            TabIndex        =   13
            Top             =   315
            Width           =   9255
            Begin VB.CheckBox chkAutoPrintBarcode 
               Appearance      =   0  'Flat
               Caption         =   "Auto Print Barcode After Request"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   300
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   2340
               Width           =   3510
            End
            Begin VB.TextBox txtExternalmsg 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   1920
               TabIndex        =   30
               TabStop         =   0   'False
               Text            =   "Unable to connect to external database"
               Top             =   3045
               Width           =   7236
            End
            Begin VB.TextBox txtSrchExtDBInterval 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   3585
               MaxLength       =   4
               TabIndex        =   29
               TabStop         =   0   'False
               Text            =   "3"
               Top             =   2655
               Width           =   750
            End
            Begin VB.CheckBox chkDisplayReqSum 
               Appearance      =   0  'Flat
               Caption         =   "Display Test Request Summary"
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   300
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   3900
               Width           =   2685
            End
            Begin VB.CheckBox chkGroupReqFrm 
               Appearance      =   0  'Flat
               Caption         =   "Group Request Form Per Section"
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   2520
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   3525
               Width           =   4215
            End
            Begin VB.CheckBox chkPrintRequestForm 
               Appearance      =   0  'Flat
               Caption         =   "Auto Print Request Form"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   300
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   3525
               Width           =   2055
            End
            Begin VB.CommandButton cmdDeptWardRoom 
               Caption         =   "Wards"
               Height          =   465
               Left            =   300
               TabIndex        =   25
               Top             =   4515
               Width           =   1545
            End
            Begin VB.CheckBox chckDisplayHospID 
               Appearance      =   0  'Flat
               Caption         =   "Display Hospital ID in Search"
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   300
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   810
               Width           =   2685
            End
            Begin VB.CheckBox chckExtDB 
               Appearance      =   0  'Flat
               Caption         =   "Use External Database"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   300
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   2685
               Width           =   2055
            End
            Begin VB.CheckBox chckEnableManualCur 
               Appearance      =   0  'Flat
               Caption         =   "Manual Start Cursor"
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   300
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   1200
               Width           =   1800
            End
            Begin VB.ComboBox cmbManualInitialCursor 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2130
               Style           =   2  'Dropdown List
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   1200
               Width           =   3105
            End
            Begin VB.TextBox txtLabelPrintCount 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   1830
               MaxLength       =   4
               TabIndex        =   18
               TabStop         =   0   'False
               Text            =   "1"
               Top             =   1995
               Width           =   750
            End
            Begin VB.CheckBox chkAutoCompAge 
               Appearance      =   0  'Flat
               Caption         =   "Auto Compute Age"
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   300
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   1605
               Width           =   2175
            End
            Begin VB.CommandButton cmdUseExtDb 
               Caption         =   "External Database Settings"
               Enabled         =   0   'False
               Height          =   435
               Left            =   5835
               TabIndex        =   16
               Top             =   7215
               Visible         =   0   'False
               Width           =   2115
            End
            Begin VB.TextBox txtSearchLimit 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1320
               MaxLength       =   4
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   420
               Width           =   750
            End
            Begin VB.CheckBox chckAutoCheckIn 
               Appearance      =   0  'Flat
               Caption         =   "Auto Check In"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   300
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   2040
               Width           =   1320
            End
            Begin VB.Label Label2 
               Caption         =   "Search Message:"
               Height          =   270
               Left            =   540
               TabIndex        =   31
               Top             =   3075
               Width           =   1305
            End
            Begin VB.Label Label1 
               Caption         =   "Search Interval:"
               Height          =   270
               Left            =   2370
               TabIndex        =   24
               Top             =   2685
               Width           =   1305
            End
            Begin VB.Label Label5 
               Caption         =   "Search Limit:"
               Height          =   270
               Left            =   300
               TabIndex        =   23
               Top             =   450
               Width           =   1305
            End
         End
         Begin VB.Frame fraDemo 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   7800
            Left            =   180
            TabIndex        =   4
            Top             =   315
            Width           =   9240
            Begin VB.TextBox txtLabel 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Enabled         =   0   'False
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
               Left            =   6345
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   2055
               Visible         =   0   'False
               Width           =   2580
            End
            Begin VB.ComboBox cmbFieldsSelection 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6375
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   1590
               Visible         =   0   'False
               Width           =   2400
            End
            Begin VB.ComboBox cmbFieldType 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "frmRequestSetting.frx":0038
               Left            =   6390
               List            =   "frmRequestSetting.frx":0045
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   1065
               Visible         =   0   'False
               Width           =   2325
            End
            Begin VB.CommandButton cmdSortUp 
               Enabled         =   0   'False
               Height          =   3465
               Left            =   75
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   705
               Width           =   555
            End
            Begin VB.CommandButton cmdSortDown 
               Enabled         =   0   'False
               Height          =   3465
               Left            =   75
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   4230
               Width           =   555
            End
            Begin VB.CheckBox chkEditSort 
               Height          =   420
               Left            =   75
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   225
               Width           =   555
            End
            Begin MSComctlLib.ImageList ImageList 
               Left            =   7575
               Top             =   7020
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               ImageWidth      =   16
               ImageHeight     =   16
               MaskColor       =   12632256
               _Version        =   393216
               BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
                  NumListImages   =   4
                  BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmRequestSetting.frx":0061
                     Key             =   "UP"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmRequestSetting.frx":0A73
                     Key             =   "Set"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmRequestSetting.frx":1485
                     Key             =   "Down"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmRequestSetting.frx":1E97
                     Key             =   ""
                  EndProperty
               EndProperty
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgSettings 
               Height          =   3030
               Left            =   5235
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   2565
               Visible         =   0   'False
               Width           =   3660
               _ExtentX        =   6456
               _ExtentY        =   5345
               _Version        =   393216
               Rows            =   1
               FixedRows       =   0
               FixedCols       =   0
               Enabled         =   0   'False
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgRequestField 
               Height          =   7485
               Left            =   690
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   210
               Width           =   8385
               _ExtentX        =   14790
               _ExtentY        =   13203
               _Version        =   393216
               Cols            =   9
               GridColorFixed  =   0
               AllowBigSelection=   0   'False
               GridLinesFixed  =   1
               ScrollBars      =   2
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
               _NumberOfBands  =   1
               _Band(0).Cols   =   9
            End
         End
      End
   End
End
Attribute VB_Name = "frmRequestSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private intRowCtr As Integer
Private intColCtr As Integer
Private intItemCtr As Integer
Private rs As New ADODB.Recordset

Private Sub chckAutoCheckIn_Click()
    If chckAutoCheckIn.Value = 1 Then
        txtLabelPrintCount.Enabled = True
    Else
        txtLabelPrintCount.Text = "1"
        txtLabelPrintCount.Enabled = False
    End If
End Sub

Private Sub chckExtDB_Click()
    If Me.chckExtDB.Value = 1 Then
        txtSrchExtDBInterval.Enabled = True
        txtExternalmsg.Enabled = True
    Else
        txtSrchExtDBInterval.Enabled = False
        txtExternalmsg.Enabled = False
    End If
End Sub

Private Sub chkEditSort_Click()
    With msfgRequestField
        If chkEditSort.Value = 1 Then
            msfgRequestField_LeaveCell
            .SelectionMode = flexSelectionByRow
            .Col = 0
            .ColSel = .Cols - 1
            .SetFocus
            cmdSortUp.Enabled = True
            cmdSortDown.Enabled = True
        Else
            .SelectionMode = flexSelectionFree
            .Col = 1
            .SetFocus
            cmdSortUp.Enabled = False
            cmdSortDown.Enabled = False
        End If
    End With
End Sub




Private Sub chkPrintRequestForm_Click()
    If chkPrintRequestForm.Value = 1 Then
        chkGroupReqFrm.Enabled = True
    Else
        chkGroupReqFrm.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeptWardRoom_Click()
   frmWardsRooms.Show 1
End Sub

Private Sub cmdSave_Click()
    msfgRequestField_LeaveCell
    With msfgRequestField
        For intRowCtr = 1 To .Rows - 1
            setGlobalSettings .TextMatrix(intRowCtr, 0), .TextMatrix(intRowCtr, 1), "RE_REQ"
            setGlobalSettings "ENA_" & .TextMatrix(intRowCtr, 0), IIf(.TextMatrix(intRowCtr, 2) = Chr(111), "0", "1"), "RE_REQ"
            setGlobalSettings "VIS_" & .TextMatrix(intRowCtr, 0), IIf(.TextMatrix(intRowCtr, 3) = Chr(111), "0", "1"), "RE_REQ"
            setGlobalSettings "MAN_" & .TextMatrix(intRowCtr, 0), IIf(.TextMatrix(intRowCtr, 4) = Chr(111), "0", "1"), "RE_REQ"
            setGlobalSettings "TYPE_" & UCase(.TextMatrix(intRowCtr, 0)), CStr(.TextMatrix(intRowCtr, 5)), "RE_REQ"
            If Len(Trim(.TextMatrix(intRowCtr, 6))) > 0 Then
                setGlobalSettings "SEL_" & UCase(.TextMatrix(intRowCtr, 0)), .TextMatrix(intRowCtr, 6), "RE_REQ"
            End If
            setGlobalSettings "SRT_" & .TextMatrix(intRowCtr, 0), .TextMatrix(intRowCtr, 7), "RE_REQ"
        Next intRowCtr
    End With
    
    setGlobalSettings "SRCH_RES_LIMIT", CStr(Val(txtSearchLimit.Text)), "RE_REQ"
    setGlobalSettings "ENA_START_CUR", CStr(chckEnableManualCur.Value), "RE_REQ"
    If chckEnableManualCur.Value = 1 Then
        setGlobalSettings "START_CUR", cmbManualInitialCursor.Text, "RE_REQ"
    End If
    setGlobalSettings "VIS_HID_SRCH", CStr(chckDisplayHospID.Value), "RE_REQ"
    setGlobalSettings "AUTO_COMP_AGE", CStr(chkAutoCompAge.Value), "RE_REQ"
    setGlobalSettings "AUTO_CHECK_IN", CStr(chckAutoCheckIn.Value), "RE_REQ"
    setGlobalSettings "AUTO_PRINT_BC", CStr(chkAutoPrintBarcode.Value), "RE_REQ"
    If chckAutoCheckIn.Value = 1 Then
        setGlobalSettings "NUM_BARCODE", txtLabelPrintCount.Text, "RE_REQ"
    Else
        setGlobalSettings "NUM_BARCODE", "1", "RE_REQ"
    End If

    SaveSetting "HostForms_RE", "ExtDemoDB\Settings", "Enable", chckExtDB.Value
    SaveSetting "HostForms_RE", "ExtDemoDB\Settings", "SearchInterval", Val(txtSrchExtDBInterval.Text)
    SaveSetting "HostForms_RE", "ExtDemoDB\Settings", "Externalmsg", txtExternalmsg.Text
    setGlobalSettings "AUTOPRNTREQFRM", CStr(chkPrintRequestForm.Value), "RE_REQ"
    setGlobalSettings "GROUPREQFRM", CStr(chkGroupReqFrm.Value), "RE_REQ"
    setGlobalSettings "DISPTESTREQSUM", CStr(chkDisplayReqSum.Value), "RE_REQ"

    Unload Me
End Sub

Private Sub cmdSortDown_Click()
    If chkEditSort.Value = 1 Then
        With msfgRequestField
            If .Row < .Rows - 1 Then
                .Col = 7
                .TextMatrix(.Row, .Col) = .TextMatrix(.Row, .Col) + 1
                .TextMatrix(.Row + 1, .Col) = .TextMatrix(.Row + 1, .Col) - 1
                .Visible = False
                .Sort = 3
                .Row = .Row + 1
                If .Row < .Rows - 2 Then
                    If .RowIsVisible(.Row + 1) = False Then
                        .TopRow = .Row
                    End If
                End If
                .Col = 0
                .ColSel = .Cols - 1
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub cmdSortUp_Click()
    If chkEditSort.Value = 1 Then
        With msfgRequestField
            If .Row > 1 Then
                .Col = 7
                .TextMatrix(.Row, .Col) = .TextMatrix(.Row, .Col) - 1
                .TextMatrix(.Row - 1, .Col) = .TextMatrix(.Row - 1, .Col) + 1
                .Visible = False
                .Sort = 3
                .Row = .Row - 1
                If .RowIsVisible(.Row) = False Then
                    .TopRow = .Row
                End If
                .Col = 0
                .ColSel = .Cols - 1
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub cmdUseExtDb_Click()
    frmDBSettings.Show 1
End Sub

Private Sub Form_Load()

    setFormColors Me
    msfgRequestField.BackColorFixed = GetSetting("LMS", "FormColors", "TitleBar", &HA66B2B)
    msfgRequestField.ForeColorFixed = GetSetting("LMS", "FormColors", "TitleText", &HFFFFFF)
    chkEditSort.Picture = ImageList.ListImages(2).Picture
    chkEditSort.DownPicture = ImageList.ListImages(4).Picture
    cmdSortUp.Picture = ImageList.ListImages(1).Picture
    cmdSortDown.Picture = ImageList.ListImages(3).Picture

    If rstSettings.State = 0 Then
        getGlobalSettings
    End If
    
    Set msfgSettings.DataSource = rstSettings
    Set msfgSettings.DataSource = Nothing
    
    setMsfgRequestFields
    
    Set rs = myConnection.Execute("EXEC sp_GetSelTypes 1,'REQ'")
    
    For intRowCtr = 0 To rs.RecordCount - 1
        cmbFieldsSelection.AddItem (rs.Fields(0))
        rs.MoveNext
    Next
    
    txtSearchLimit.Text = FindSetting2("SRCH_RES_LIMIT", msfgSettings, "0")
    chckEnableManualCur.Value = FindSetting2("ENA_START_CUR", msfgSettings, "0")
    
    If chckEnableManualCur.Value = 1 Then
        Dim tempStartCurs As String
                
        tempStartCurs = FindSetting("START_CUR", "")
        For intRowCtr = 0 To cmbManualInitialCursor.ListCount - 1
            If cmbManualInitialCursor.List(intRowCtr) = tempStartCurs Then
                cmbManualInitialCursor.ListIndex = intRowCtr
                Exit For
            End If
        Next
    End If
    
    chckDisplayHospID.Value = FindSetting2("VIS_HID_SRCH", msfgSettings, "0")
    chkAutoCompAge.Value = FindSetting2("AUTO_COMP_AGE", msfgSettings, "0")
    chckAutoCheckIn.Value = FindSetting2("AUTO_CHECK_IN", msfgSettings, "1")
    chkAutoPrintBarcode.Value = FindSetting2("AUTO_PRINT_BC", msfgSettings, "1")
    
    If chckAutoCheckIn.Value = 1 Then
        txtLabelPrintCount.Text = FindSetting2("NUM_BARCODE", msfgSettings, "1")
    End If
    
    chckExtDB.Value = Val(GetSetting("HostForms_RE", "ExtDemoDB\Settings", "Enable", "0"))
    txtSrchExtDBInterval.Text = Val(GetSetting("HostForms_RE", "ExtDemoDB\Settings", "SearchInterval", "3"))
    txtExternalmsg.Text = GetSetting("HostForms_RE", "ExtDemoDB\Settings", "Externalmsg", "Unable to connect to external database")
    
    
    chkPrintRequestForm.Value = FindSetting2("AUTOPRNTREQFRM", msfgSettings, "0")
    chkGroupReqFrm.Value = FindSetting2("GROUPREQFRM", msfgSettings, "0")

    chkDisplayReqSum.Value = FindSetting2("DISPTESTREQSUM", msfgSettings, "0")
    
    SSTab.Tab = 0

End Sub

Private Sub setMsfgRequestFields()

    Dim rsRequestFields
        
    strSQLStatement = "EXEC sp_RequestIniSettings"
    Set rsRequestFields = myConnection.Execute(strSQLStatement)
    With msfgRequestField
        .ColWidth(0) = 1500
        .ColWidth(1) = 1500
        .ColWidth(2) = 600
        .ColWidth(3) = 600
        .ColWidth(4) = 850
        .ColWidth(5) = 1000
        .ColWidth(6) = 1500
        .ColWidth(7) = 500
        .ColWidth(.Cols - 1) = 0
        
        .TextArray(1) = "Label"
        .TextArray(2) = "Enable"
        .TextArray(3) = "Visible"
        .TextArray(4) = "Mandatory"
        .TextArray(5) = "Type"
        .TextArray(6) = "Data Field"
        .TextArray(7) = "Sort"
        .TextArray(.Cols - 1) = "InitSet"
        
        For intRowCtr = 0 To rsRequestFields.Fields.Count - 1
            .Row = .Rows - 1
            .Col = .Cols - 1
            .TextMatrix(.Row, .Col) = rsRequestFields.Fields(intRowCtr).Value
            .Col = 0
            .TextMatrix(.Row, .Col) = rsRequestFields.Fields(intRowCtr).Name
            .CellAlignment = 0
            .Col = 1
            .TextMatrix(.Row, .Col) = FindSetting2(UCase(.TextMatrix(.Row, 0)), msfgSettings, .TextMatrix(.Row, 0))
            .CellAlignment = 0
            .Col = 2
            .CellFontName = "Wingdings"
            .CellFontBold = True
            .CellFontSize = 12
            .CellAlignment = 4
            Select Case Mid(.TextMatrix(.Row, .Cols - 1), 1, 1)
                Case "1"
                    .TextMatrix(.Row, .Col) = IIf(FindSetting2("ENA_" & UCase(.TextMatrix(.Row, 0)), msfgSettings, "0") = "0", Chr(111), Chr(254))
                Case "2"
                    .TextMatrix(.Row, .Col) = Chr(254)
                Case "3"
                    .TextMatrix(.Row, .Col) = Chr(111)
            End Select
            
            .Col = 3
            .CellFontName = "Wingdings"
            .CellFontBold = True
            .CellFontSize = 12
            .CellAlignment = 4
            Select Case Mid(.TextMatrix(.Row, .Cols - 1), 2, 1)
                Case "1"
                    .TextMatrix(.Row, .Col) = IIf(FindSetting2("VIS_" & UCase(.TextMatrix(.Row, 0)), msfgSettings, "0") = "0", Chr(111), Chr(254))
                Case "2"
                    .TextMatrix(.Row, .Col) = Chr(254)
                Case "3"
                    .TextMatrix(.Row, .Col) = Chr(111)
            End Select
            
            .Col = 4
            .CellFontName = "Wingdings"
            .CellFontBold = True
            .CellFontSize = 12
            .CellAlignment = 4
            Select Case Mid(.TextMatrix(.Row, .Cols - 1), 3, 1)
                Case "1"
                    .TextMatrix(.Row, .Col) = IIf(FindSetting2("MAN_" & UCase(.TextMatrix(.Row, 0)), msfgSettings, "0") = "0", Chr(111), Chr(254))
                Case "2"
                    .TextMatrix(.Row, .Col) = Chr(254)
                Case "3"
                    .TextMatrix(.Row, .Col) = Chr(111)
            End Select
            
            Select Case Mid(.TextMatrix(.Row, .Cols - 1), 4, 1)
                Case "1"
                    .TextMatrix(.Row, 5) = UCase(FindSetting2("TYPE_" & UCase(.TextMatrix(.Row, 0)), msfgSettings, "TEXT"))
                Case "4"
                    .TextMatrix(.Row, 5) = "TEXT"
                Case "5"
                    .TextMatrix(.Row, 5) = "SELECTION"
            End Select
            
            .TextMatrix(.Row, 6) = FindSetting2("SEL_" & UCase(.TextMatrix(.Row, 0)), msfgSettings, "")
            .TextMatrix(.Row, 7) = FindSetting2("SRT_" & UCase(.TextMatrix(.Row, 0)), msfgSettings, .Row)
            .RowHeight(.Row) = 300
            .Rows = .Rows + 1
        Next
        .Rows = .Rows - 1
        .Col = 7
        .Sort = 1
        .Col = 1
        .Row = 1
    End With
           
End Sub

Private Sub msfgRequestField_Click()
    If chkEditSort.Value = 0 Then
        With msfgRequestField
            txtLabel.Enabled = False
            txtLabel.Visible = False
            cmbFieldType.Visible = False
            cmbFieldType.Enabled = False
            cmbFieldsSelection.Visible = False
            cmbFieldsSelection.Enabled = False
            Select Case .Col
                Case 1
                    txtLabel.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
                    txtLabel.Text = .TextMatrix(.Row, .Col)
                    txtLabel.Visible = True
                    txtLabel.Enabled = True
                    txtLabel.SetFocus
                Case 2
                    If Mid(.TextMatrix(.Row, .Cols - 1), 1, 1) = "1" Then
                        .CellFontName = "Wingdings"
                        .CellFontBold = True
                        .CellFontSize = 12
                        .CellAlignment = 4
                        .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = Chr(111), Chr(254), Chr(111))
                    Else
                        MsgBox "This field is defined as default.", vbInformation + vbOKOnly, "LMS"
                    End If
                Case 3
                     If Mid(.TextMatrix(.Row, .Cols - 1), 2, 1) = "1" Then
                        .CellFontName = "Wingdings"
                        .CellFontBold = True
                        .CellFontSize = 12
                        .CellAlignment = 4
                        .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = Chr(111), Chr(254), Chr(111))
                     Else
                        MsgBox "This field is defined as default.", vbInformation + vbOKOnly, "LMS"
                    End If
                Case 4
                    If Mid(.TextMatrix(.Row, .Cols - 1), 3, 1) = "1" Then
                        .CellFontName = "Wingdings"
                        .CellFontBold = True
                        .CellFontSize = 12
                        .CellAlignment = 4
                        .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "o", "þ", "o")
                     Else
                        MsgBox "This field is defined as default.", vbInformation + vbOKOnly, "LMS"
                    End If
                Case 5
                    If Mid(.TextMatrix(.Row, .Cols - 1), 4, 1) = "1" Then
                        cmbFieldType.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
                        If Len(.TextMatrix(.Row, .Col)) > 0 Then
                            For intRowCtr = 0 To cmbFieldType.ListCount - 1
                                If .TextMatrix(.Row, .Col) = cmbFieldType.List(intRowCtr) Then
                                    cmbFieldType.ListIndex = intRowCtr
                                    Exit For
                                End If
                            Next
                        Else
                            cmbFieldType.ListIndex = -1
                        End If
                        cmbFieldType.Visible = True
                        cmbFieldType.Enabled = True
                        cmbFieldType.SetFocus
                    Else
                        MsgBox "This field is defined as default.", vbInformation + vbOKOnly, "LMS"
                    End If
                Case 6
                    If UCase(.TextMatrix(.Row, 5)) = "COMBO" Then
                        cmbFieldsSelection.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
                        If Len(.TextMatrix(.Row, .Col)) > 0 Then
                            For intRowCtr = 0 To cmbFieldsSelection.ListCount - 1
                                If .TextMatrix(.Row, .Col) = cmbFieldsSelection.List(intRowCtr) Then
                                    cmbFieldsSelection.ListIndex = intRowCtr
                                    Exit For
                                End If
                            Next
                        Else
                            cmbFieldsSelection.ListIndex = -1
                        End If
                        cmbFieldsSelection.Visible = True
                        cmbFieldsSelection.Enabled = True
                        cmbFieldsSelection.SetFocus
                    ElseIf UCase(.TextMatrix(.Row, 5)) = "SELECTION" Then
                        txtLabel.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
                        If Len(.TextMatrix(.Row, .Col)) > 0 Then
                            txtLabel.Text = .TextMatrix(.Row, .Col)
                            txtLabel.SelStart = Len(txtLabel.Text)
                        Else
                            txtLabel.Text = ""
                        End If
                        txtLabel.Visible = True
                        txtLabel.Enabled = True
                        txtLabel.SetFocus
                    End If
            End Select
        End With
    Else
        msfgRequestField.Col = 0
        msfgRequestField.ColSel = msfgRequestField.Cols - 1
    End If
End Sub

Private Sub msfgRequestField_LeaveCell()
    If txtLabel.Visible Or cmbFieldType.Visible Or cmbFieldsSelection.Visible = True Then
        With msfgRequestField
            If .Col = 1 Then
                .TextMatrix(.Row, .Col) = txtLabel.Text
                txtLabel.Text = ""
                txtLabel.Visible = False
                txtLabel.Enabled = False
            ElseIf .Col = 5 Then
                If Mid(.TextMatrix(.Row, .Cols - 1), 4, 1) = "1" Then
                    .TextMatrix(.Row, .Col) = cmbFieldType.Text
                    cmbFieldType.ListIndex = -1
                    cmbFieldType.Visible = False
                    cmbFieldType.Enabled = False
                End If
            ElseIf .Col = 6 Then
                If UCase(.TextMatrix(.Row, 5)) = "COMBO" Then
                    .TextMatrix(.Row, .Col) = cmbFieldsSelection.Text
                    cmbFieldsSelection.ListIndex = -1
                    cmbFieldsSelection.Visible = False
                    cmbFieldsSelection.Enabled = False
                ElseIf UCase(.TextMatrix(.Row, 5)) = "SELECTION" Then
                    .TextMatrix(.Row, .Col) = txtLabel.Text
                    txtLabel.Text = ""
                    txtLabel.Visible = False
                    txtLabel.Enabled = False
                End If
            End If
        End With
    End If
End Sub

Private Sub msfgRequestField_Scroll()
    msfgRequestField_LeaveCell
End Sub


Private Sub txtLabel_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtLabel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then
        If msfgRequestField.Row > 1 Then
            msfgRequestField_LeaveCell
            msfgRequestField.Row = msfgRequestField.Row - 1
            msfgRequestField.SetFocus
            msfgRequestField_Click
            highLight
        End If
    ElseIf KeyCode = 40 Then
        If msfgRequestField.Row < msfgRequestField.Rows - 1 Then
            msfgRequestField_LeaveCell
            msfgRequestField.Row = msfgRequestField.Row + 1
            msfgRequestField.SetFocus
            msfgRequestField_Click
            highLight
        End If
    End If
End Sub

Private Sub txtLabel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If msfgRequestField.Row < msfgRequestField.Rows - 1 Then
              msfgRequestField_LeaveCell
              msfgRequestField.Row = msfgRequestField.Row + 1
              msfgRequestField.SetFocus
              msfgRequestField_Click
              highLight
          End If
          KeyAscii = 0
    End If
End Sub

Private Sub txtLabelPrintCount_GotFocus()
    highLight
End Sub

Private Sub txtLabelPrintCount_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48 And KeyAscii <= 56) Or KeyAscii = 8 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtReqLabel_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtSearchLimit_GotFocus()
    highLight
End Sub

Private Sub txtSearchLimit_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48 And KeyAscii <= 56) Or KeyAscii = 8 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
