VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmEditRequest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Request"
   ClientHeight    =   10320
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   16575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   16575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   855
      Left            =   19770
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6975
      Width           =   1350
   End
   Begin MSComctlLib.Toolbar tlbarReq 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   1058
      ButtonWidth     =   2858
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imgListToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Print Label"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Remove Test"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Print Request Form"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Close"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   9750
      Index           =   0
      Left            =   75
      TabIndex        =   5
      Top             =   510
      Width           =   16440
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgExamDel 
         Height          =   1170
         Left            =   4725
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   5310
         Visible         =   0   'False
         Width           =   11550
         _ExtentX        =   20373
         _ExtentY        =   2064
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         Enabled         =   0   'False
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   120
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   8310
         Width           =   2145
      End
      Begin VB.CheckBox chkStat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4800
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   920
         Width           =   195
      End
      Begin VB.CommandButton cmdExamMenu 
         Appearance      =   0  'Flat
         Caption         =   "Exam List"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   14820
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   1365
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgExamDet 
         Height          =   1170
         Left            =   4725
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   6510
         Visible         =   0   'False
         Width           =   11550
         _ExtentX        =   20373
         _ExtentY        =   2064
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         Enabled         =   0   'False
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin VB.VScrollBar VScrollExam 
         Enabled         =   0   'False
         Height          =   2370
         Left            =   16050
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   870
         Width           =   255
      End
      Begin VB.VScrollBar VScrollDemo 
         Enabled         =   0   'False
         Height          =   2370
         Left            =   4305
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   855
         Width           =   255
      End
      Begin VB.TextBox txtmsfgExam 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5910
         MaxLength       =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   8865
      End
      Begin VB.ComboBox cmbmsfgResult 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         ItemData        =   "frmEditRequest.frx":0000
         Left            =   1530
         List            =   "frmEditRequest.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   4530
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.TextBox txtmsfgResult 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Left            =   1200
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2910
         Visible         =   0   'False
         Width           =   1245
      End
      Begin MSComctlLib.ImageList imglstIcon 
         Left            =   15645
         Top             =   4620
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":0004
               Key             =   "newrequest"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":03FE
               Key             =   "checkin"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":0822
               Key             =   "ongoing"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":0C47
               Key             =   "ready"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":104B
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":1453
               Key             =   "new"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":184F
               Key             =   "printed"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":1C6E
               Key             =   "saved"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":2081
               Key             =   "unapprove"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgListToolBar 
         Left            =   17430
         Top             =   5670
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   15
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":2473
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":2864
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":2C70
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":3083
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":3487
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":384F
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":3CA1
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":406E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":4405
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":479B
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":4B68
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":4F76
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":532F
               Key             =   "preview"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":570E
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditRequest.frx":5B12
               Key             =   "Delete"
            EndProperty
         EndProperty
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   6120
         Left            =   45
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   840
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   10795
         _Version        =   393216
         Rows            =   1
         Cols            =   12
         FixedRows       =   0
         GridColorFixed  =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
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
         _Band(0).Cols   =   12
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgExam 
         Height          =   4425
         Left            =   4725
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   840
         Width           =   11610
         _ExtentX        =   20479
         _ExtentY        =   7805
         _Version        =   393216
         GridColor       =   -2147483647
         GridColorFixed  =   -2147483640
         GridColorUnpopulated=   16777215
         FocusRect       =   0
         HighLight       =   2
         GridLinesFixed  =   1
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
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   90
         TabIndex        =   18
         Top             =   7890
         Width           =   2265
      End
      Begin VB.Label lblLocked 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmEditRequest.frx":6524
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1725
         Left            =   90
         TabIndex        =   16
         Top             =   7950
         Visible         =   0   'False
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Code/ Name"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   3
         Left            =   4680
         TabIndex        =   7
         Top             =   195
         Width           =   2415
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Demographics"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   0
         Left            =   75
         TabIndex        =   6
         Top             =   405
         Width           =   2265
      End
   End
   Begin VB.Frame fraTitle 
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   15405
      TabIndex        =   4
      Top             =   2565
      Width           =   825
   End
   Begin VB.Menu mnuCommands 
      Caption         =   "Commands"
      Begin VB.Menu mnucmdSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnucmdPrntLbl 
         Caption         =   "Print Label"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu mnucmdRmvTest 
         Caption         =   "Remove Test"
         Enabled         =   0   'False
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnucmdPrntReqFrm 
         Caption         =   "Print Request Form"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnucmdCnlReq 
         Caption         =   "Cancel Request"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnucmdClose 
         Caption         =   "Close"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmEditRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GWL_STYLE = (-16)
Private Const WS_VSCROLL = &H200000

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private wndStyle As Long


Private blnNewPatient As Boolean
Private edtPatientID As String
Private edtPatientHospID As String
Private edtSpecimen As String
Private edtRequestID As String
Private edtSectionID As String
Private edtSampleNumber As String
Private edtUserID As String
Private edtUserName As String
Private newPatientName As String
Private blnWithSettingsDemo As Boolean
Private blnWithSettingsReq As Boolean
Private blnWithSettings As Boolean
Private blnNoExistingTest As Boolean
Private blnSaveButtonClick As Boolean
Private intRowReqAge As Integer 'Temp Variable to locate Row num of req age
Private currSampleStatus As String

Private initmsfgExamRowNum As Integer

Private strUserRights As String

Private ENA_START_CUR As String
Private START_CUR As String
Private AUTO_CHECK_IN As String
Private LABEL_NUM_COPY As String
Private AUTO_COMP_AGE As String
Private AUTOPRNTREQFRM As String
Private AUTO_PRINT_BC  As String
Private GROUPREQFRM As String
Private DISPTESTREQSUM

Private blnFormIsLocked As Boolean


Private Sub chkStat_Click()
    Dim ctrmsfgexam As Integer
    msfgResult_LeaveCell
    If chkStat.Value = 1 Then
        For ctrmsfgexam = 1 To msfgExam.Rows - 1
            With msfgExam
                If Len(.TextMatrix(ctrmsfgexam, 1)) > 0 Then
                    .TextMatrix(ctrmsfgexam, 11) = Chr(254)
                End If
            End With
        Next
    Else
       For ctrmsfgexam = 1 To msfgExam.Rows - 1
            With msfgExam
                If Len(.TextMatrix(ctrmsfgexam, 1)) > 0 Then
                    .TextMatrix(ctrmsfgexam, 11) = Chr(111)
                End If
            End With
        Next
    End If
End Sub

''''''''''''''''''''''''''''''''
'Settings DEMO and REQ GRID
'COL00=Lablel Display
'COL01=DB Value
'COL02=DB Column Label
'COL03=Field Enable
'COL04=Field Visible
'COL05=Field Mandatory
'COL06=Field Type
'COL07=IF Field Type Is COMBO, DB SelectionValues Type
'COL08=
'COL09=Manual Start Cursor
'COL10=DB Id Value
'COL11=Request Value Indicator
''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''
'DB Value Location DEMO
'ROW00-PID
'ROW01-PHOSTPID
'ROW02-PLNAME
'ROW03-PFNAME
'ROW04-PMNAME
'ROW05-PDOB
'ROW06-PGENDER
'ROW07-PCIVILSTATUS
'ROW08-PPHICID
'ROW09-PADDRESS1
'ROW10-PADDRESS2
'ROW11-PEMAIL
'ROW12-PPHONE
''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''
'DB Value Location REQ
'ROW13-RequestAge
'ROW14-RequestPID
'ROW15-RequestID
'ROW16-RequestStatus
'ROW17-RequestDate
'ROW18-RequestBy
'ROW19-RequestHospReqID
'ROW20-RequestWard
'ROW21-RequestRoom
'ROW22-RequestBed
'ROW23-ReqRate
'ROW24-RequestLocation
'ROW25-RequestPriority
'ROW26-RequestPType
'ROW27-RequestPhysicianID
'ROW28-Reserve1
'ROW29-Reserve2
'ROW30-Reserve3
'ROW31-Reserve4
'ROW32-Reserve5
'ROW33-RequestComment
''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''
'Setting of msfgexam
'COL00 = status(icon)
'COL01 = exam code
'COL02 = exam name
'COL03 = specimen
'COL04 = section name
'COL05 = amount
'COL06 = status
'COL07 = section id
'COL08 = sampleid
'COL09 = rvid
'COL10 = labexamtype
'COL11 = Priority
'COL12 = PanelID
'COL13 = PriorityOrig
''''''''''''''''''''''''''''''''

Private Sub cmbmsfgResult_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim msfgResultCurrentRow As Integer
    Dim msfgResultNewRow As Integer

    If msfgResult.Col = 1 Then
        If KeyCode = vbKeyUp Then
            If msfgResult.Row > 0 Then
                msfgResultCurrentRow = msfgResult.Row
                msfgResult_LeaveCell

                For msfgResultNewRow = msfgResultCurrentRow - 1 To 0 Step -1
                    If msfgResult.TextMatrix(msfgResultNewRow, 3) = "1" And _
                        msfgResult.TextMatrix(msfgResultNewRow, 4) = "1" Then
                        msfgResult.Row = msfgResultNewRow
                        Exit For
                    End If
                Next

                msfgResult.SetFocus
                msfgResult_EnterCell
                highLight
            Else
                cmbmsfgResult.SetFocus
            End If

        ElseIf KeyCode = vbKeyDown Or KeyCode = 9 Then
            If msfgResult.Row < msfgResult.Rows - 1 Then
                msfgResultCurrentRow = msfgResult.Row
                msfgResult_LeaveCell
    
                For msfgResultNewRow = msfgResult.Row + 1 To msfgResult.Rows - 1
                    If msfgResult.TextMatrix(msfgResultNewRow, 3) = "1" And _
                    msfgResult.TextMatrix(msfgResultNewRow, 4) = "1" Then
                        msfgResult.Row = msfgResultNewRow
                        Exit For
                    End If
                Next
    
                msfgResult.SetFocus
                msfgResult_EnterCell
                highLight
            Else
                cmbmsfgResult.SetFocus
            End If
        
        ElseIf KeyCode = 46 Or KeyCode = 8 Then
            cmbmsfgResult.ListIndex = -1
        End If
        
    End If
End Sub

Private Sub cmbmsfgResult_KeyPress(KeyAscii As Integer)

    Dim msfgResultCurrentRow As Integer
    Dim msfgResultNewRow As Integer
    
    If KeyAscii = 13 Then
        msfgResultCurrentRow = msfgResult.Row
        msfgResult_LeaveCell
        
        With msfgResult
            For msfgResultNewRow = msfgResult.Row + 1 To msfgResult.Rows - 1
                If msfgResult.TextMatrix(msfgResultNewRow, 3) = "1" And msfgResult.TextMatrix(msfgResultNewRow, 4) = "1" Then
                    msfgResult.Row = msfgResultNewRow
                    Exit For
                End If
            Next
        End With
        
        If msfgResult.Row = msfgResultCurrentRow Then
            msfgResult_EnterCell
'            If cmdSaveDemo.Enabled Then
'                cmdSaveDemo.SetFocus
'                txtmsfgResult.Visible = False
'                txtmsfgResult.Enabled = False
'            End If
        Else
'            msfgResult.SetFocus
            msfgResult_EnterCell
            highLight
        End If
        KeyAscii = 0
    ElseIf KeyAscii = 9 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExamMenu_Click()
    Dim strSelectionValue As String
    Dim intColClear As Integer
    Dim KD_arryExamDetFromRemoved
    Dim KD_arryExamSelected() As String
    
    If Not blnNoExistingTest Then
        strSelectionValue = showfrmExamSelection(txtmsfgExam.Text, edtUserID, edtSpecimen, edtSectionID)
    Else
        strSelectionValue = showfrmExamSelection(txtmsfgExam.Text, edtUserID, "", "")
    End If
    
    If isExiting Then
        Unload Me
        Exit Sub
    End If
    
    If Len(strSelectionValue) > 0 Then
        KD_arryExamSelected = Split(strSelectionValue, "|")
        If Not blnExamOccured(strSelectionValue) Then
            With msfgExam
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = "0"
                .CellAlignment = 1
                .TextMatrix(.Row, 1) = KD_arryExamSelected(0)
                .TextMatrix(.Row, 2) = KD_arryExamSelected(1)
                .TextMatrix(.Row, 3) = KD_arryExamSelected(2)
                .TextMatrix(.Row, 4) = KD_arryExamSelected(5)
                .TextMatrix(.Row, 5) = ""
                .TextMatrix(.Row, 6) = "0"
                .TextMatrix(.Row, 7) = KD_arryExamSelected(4)
                .TextMatrix(.Row, 8) = ""
                .TextMatrix(.Row, 9) = ""
                .TextMatrix(.Row, 10) = KD_arryExamSelected(3)
                .Col = 11
                .CellFontName = "Wingdings"
                .CellFontBold = True
                .CellFontSize = 12
                .CellAlignment = 4
                .TextMatrix(.Row, 11) = IIf(chkStat.Value = 0, Chr(111), Chr(254))
                
                If Len(edtRequestID) > 0 Then
                    .TextMatrix(.Row, 8) = edtSampleNumber
                    KD_arryExamDetFromRemoved = Split(RemoveFromTestToDel(KD_arryExamSelected(0)), "|")
                    If UBound(KD_arryExamDetFromRemoved) > 0 Then
                        .TextMatrix(.Row, 0) = KD_arryExamDetFromRemoved(0)
                        .TextMatrix(.Row, 6) = KD_arryExamDetFromRemoved(0)
                        .TextMatrix(.Row, 9) = KD_arryExamDetFromRemoved(1)
                        .TextMatrix(.Row, 11) = KD_arryExamDetFromRemoved(2)
                        .TextMatrix(.Row, 12) = KD_arryExamDetFromRemoved(3)
                        .Col = 0
                        .CellPictureAlignment = 4
                        Set .CellPicture = imglstIcon.ListImages(2).Picture
                    End If
                End If
                
                
                .Col = 0
                .CellForeColor = vbWhite
            
                AddExamDet .Row
                
                If Len(.TextMatrix(.Rows - 1, 0)) > 0 Then
                    .Rows = .Rows + 1
                End If
                .Row = .Rows - 1

                .Col = 0
                .CellForeColor = vbWhite
                .CellBackColor = vbWhite
                Set .CellPicture = imglstIcon.ListImages(6).Picture
                .CellPictureAlignment = 4
                .Col = 11
                .CellFontName = "Wingdings"
                .CellFontBold = True
                .CellFontSize = 12
                .TextMatrix(.Row, .Col) = Chr(111)
                .CellAlignment = 4
                .Col = 1
                
                .RowHeight(.Row) = 350
                txtmsfgExam.Text = ""
                .CellAlignment = 0

                If .Rows > 20 Then
                    VScrollExam.Visible = False
                Else
                    VScrollExam.Visible = True
                End If
                
                .TopRow = .Rows - 1
            End With
        Else
            MsgBox "Unable to Add " & KD_arryExamSelected(1), vbInformation + vbOKOnly, "LMS"
        End If
    End If
    txtmsfgExam.SetFocus
End Sub

Private Sub Form_Load()
    On Error Resume Next
    setFormColors Me
    msfgExam.BackColorFixed = GetSetting("LMS", "FormColors", "TitleBar", &HA66B2B)
    msfgExam.ForeColorFixed = GetSetting("LMS", "FormColors", "TitleText", &HFFFFFF)
    
    If rstSettings.State = 0 Then
        getGlobalSettings
    End If
    
    ENA_START_CUR = FindSetting("ENA_START_CUR", "0")
    START_CUR = FindSetting("START_CUR", "")

    AUTO_CHECK_IN = FindSetting("AUTO_CHECK_IN", "1")
    LABEL_NUM_COPY = GetSetting("LMS", "Barcode", "Copy", "1")
    
    AUTO_COMP_AGE = FindSetting("AUTO_COMP_AGE", "0")
    
    AUTO_PRINT_BC = FindSetting("AUTO_PRINT_BC", "1")
    AUTOPRNTREQFRM = FindSetting("AUTOPRNTREQFRM", "0")
    GROUPREQFRM = FindSetting("GROUPREQFRM", "0")
    
    DISPTESTREQSUM = FindSetting("DISPTESTREQSUM", "0")
    
    strUserRights = getValue(getUserDet(edtUserID), 0)
    
    setmsfgResultDemo
    setmsfgResultReq
    msfgResult.Height = 7000
    msfgExam.Height = 7000
    setmsfgExam
      
    cmdExamMenu.Picture = imglstIcon.ListImages(1).Picture
    
    lblText(3).Caption = "Exam " & vbCr & "Code/Name"
    
    If Not blnNewPatient Then
        If ReqIsLocked Then
            blnFormIsLocked = lockForm("Locked. Sample is already opened by another user")
        Else
            ReqLock
            blnFormIsLocked = False
        End If
    End If
  
    initmsfgExamRowNum = msfgExam.Rows
    
    If Not blnFormIsLocked Then
        setIntialCursor
            
        If blnWithSettingsDemo = True And blnWithSettingsReq = True Then
            blnWithSettings = True
        Else
            blnWithSettings = False
        End If
    
        tlbarReq.Buttons(1).Enabled = blnWithSettings
        mnucmdSave.Enabled = blnWithSettings
        
        If blnNewPatient Then
            tlbarReq.Buttons(2).Enabled = False
            tlbarReq.Buttons(4).Enabled = False
            mnucmdPrntLbl.Enabled = False
            mnucmdPrntReqFrm.Enabled = False
        Else
            tlbarReq.Buttons(2).Enabled = IIf((Len(Trim(edtSampleNumber)) > 0), True, False)
            mnucmdPrntLbl.Enabled = IIf((Len(Trim(edtSampleNumber)) > 0), True, False)
            tlbarReq.Buttons(4).Enabled = tlbarReq.Buttons(2).Enabled
            mnucmdPrntReqFrm.Enabled = mnucmdPrntLbl.Enabled
        End If
        
        msfgResult.Enabled = blnWithSettings
        msfgExam.Enabled = blnWithSettings
        
        If strUserRights = 4 Or strUserRights = 1 Or strUserRights = 2 Then
            tlbarReq.Buttons(3).Visible = True
        Else
            tlbarReq.Buttons(3).Enabled = False
            tlbarReq.Buttons(3).Visible = False
            mnucmdRmvTest.Enabled = False
            mnucmdRmvTest.Visible = False
        End If
        
    End If
    
    If Me.Width > Screen.Width Then
        sizeFix
    End If
    VScrollDemo.Height = msfgResult.Height - 10
    wndStyle = GetWindowLong(msfgResult.hwnd, GWL_STYLE)
    If (wndStyle And WS_VSCROLL) <> 0 Then
        VScrollDemo.Visible = False
    End If
    VScrollExam.Height = msfgExam.Height - 50
    wndStyle = GetWindowLong(msfgExam.hwnd, GWL_STYLE)
    If (wndStyle And WS_VSCROLL) <> 0 Then
        VScrollExam.Visible = False
    End If

    chkStat.Left = 14950
    chkStat.ForeColor = msfgExam.BackColorFixed
    chkStat.BackColor = msfgExam.BackColorFixed
    
    With txtRemarks
        .Top = 8310
        .Left = 120
        .Width = 16065
        .Height = 1275
    End With
End Sub

Private Sub setmsfgResultDemo()
    Dim intRowCtrDemo As Integer
    Dim intColCtrDemo As Integer
    Dim rsedtRequestDemo As New ADODB.Recordset
    On Error GoTo errfound
    With msfgResult
        'set design and label of grid
        
        If blnNewPatient Then
            strSQLStatement = "EXEC sp_PatientDemo_select 'NEWPATIENT'"
            Set rsedtRequestDemo = myConnection.Execute(strSQLStatement)
        Else
            If Len(edtPatientHospID) = 0 Then
                strSQLStatement = "EXEC sp_PatientDemo_select '" & edtPatientID & "'"
                Set rsedtRequestDemo = myConnection.Execute(strSQLStatement)
            Else
                strSQLStatement = "EXEC sp_PatientDemoTemp_select " & intWorkstation & ",'" & edtPatientHospID & "'"
                Set rsedtRequestDemo = myConnection.Execute(strSQLStatement)
                If IsNull(rsedtRequestDemo.Fields(0)) Or Len(Trim(rsedtRequestDemo.Fields(0))) = 0 Then
                    blnNewPatient = True
                Else
                    blnNewPatient = False
                End If
            End If
        End If
        
        .Rows = rsedtRequestDemo.Fields.Count
        
        For intRowCtrDemo = 0 To .Rows - 1
            .Col = 0
            .Row = intRowCtrDemo
            
            .TextMatrix(.Row, 2) = rsedtRequestDemo.Fields(.Row).Name
            .TextMatrix(.Row, 0) = FindSetting(.TextMatrix(.Row, 2), .TextMatrix(.Row, 2))
            .TextMatrix(.Row, 3) = FindSetting("ENA_" & .TextMatrix(.Row, 2), "0")
            .TextMatrix(.Row, 4) = FindSetting("VIS_" & .TextMatrix(.Row, 2), "1")
            .TextMatrix(.Row, 5) = FindSetting("MAN_" & .TextMatrix(.Row, 2), "1")
            .TextMatrix(.Row, 6) = FindSetting("TYPE_" & .TextMatrix(.Row, 2), "TEXT")
            .TextMatrix(.Row, 7) = IIf(.TextMatrix(.Row, 6) = "COMBO", FindSetting("SEL_" & .TextMatrix(.Row, 2), ""), "")
            
            If blnWithSettingsDemo = False And .TextMatrix(.Row, 3) = "1" Then
                blnWithSettingsDemo = True
            End If

            If Len(START_CUR) > 0 Then
                If .TextMatrix(.Row, 0) = START_CUR And .TextMatrix(.Row, 3) = "1" And .TextMatrix(.Row, 4) = "1" Then
                    .TextMatrix(.Row, 9) = "1"
                End If
            End If
    
            .Col = 1
            .CellAlignment = 1
        
            .CellBackColor = IIf(.TextMatrix(.Row, 3) = "1", &H80000005, &H80000011)
            
            If blnNewPatient Then
                If rsedtRequestDemo.RecordCount > 0 Then
                    .TextMatrix(.Row, 1) = IIf(IsNull(rsedtRequestDemo.Fields(.Row).Value), "", Trim(rsedtRequestDemo.Fields(.Row).Value))
                Else
                    If Len(newPatientName) > 0 Then
                        If .Row = 2 Then
                            If InStr(1, newPatientName, ",") > 0 Then
                                .TextMatrix(.Row, 1) = Trim(Mid(newPatientName, 1, InStr(newPatientName, ",") - 1))
                            Else
                                .TextMatrix(.Row, 1) = newPatientName
                            End If
                        ElseIf .Row = 3 Then
                            .TextMatrix(.Row, 1) = IIf(InStr(1, newPatientName, ",") > 0, Trim(Mid(newPatientName, InStr(newPatientName, ",") + 1)), "")
                        Else
                            
                        End If
                    Else
                        .TextMatrix(.Row, 1) = ""
                    End If
                End If
                .Col = 11
                .CellAlignment = 1
                .TextMatrix(.Row, 11) = "NEWPATIENT"
            Else
                If rsedtRequestDemo.RecordCount > 0 Then
                    .TextMatrix(.Row, 1) = IIf(IsNull(rsedtRequestDemo.Fields(.Row).Value), "", Trim(rsedtRequestDemo.Fields(.Row).Value))
                    .Col = 11
                    .CellAlignment = 1
                    .TextMatrix(.Row, 11) = IIf(IsNull(rsedtRequestDemo.Fields(.Row).Value), "", Trim(rsedtRequestDemo.Fields(.Row).Value))
                End If
            End If

            If .TextMatrix(.Row, 2) = "DOB" Then
                If Len(.TextMatrix(.Row, 1)) > 0 Then
                    .TextMatrix(.Row, 1) = Format(.TextMatrix(.Row, 1), "MM/dd/yyyy")
                End If
            End If
            
            .RowHeight(.Row) = IIf(.TextMatrix(.Row, 4) = "1", 350, 0)
        Next

        For intColCtrDemo = 0 To .Cols - 1
            .ColWidth(intColCtrDemo) = 0
        Next

        .ColWidth(0) = 1500
        .ColWidth(1) = 2750

    End With
errfound:
    If Err.Number <> 0 Then MsgBox "setmsfgResultDemo: " & Err.Description, vbOKOnly + vbExclamation, App.Title
End Sub

Private Sub setmsfgResultReq()

    Dim numRowDemo As Integer
    
    Dim intRowCtrReq As Integer
    Dim intColCtrReq As Integer
    
    Dim rsedtRequestReq As New ADODB.Recordset
    
    Dim ENA_SCOL_REQD_EDT As String
    Dim strStartRowEdtReq As String

    ENA_SCOL_REQD_EDT = FindSetting("ENA_SCOL_REQD_EDT", "0")
    
    If ENA_SCOL_REQD_EDT = "1" Then
        strStartRowEdtReq = FindSetting("CUR_REQD_EDT", "")
    End If
    
    With msfgResult
        
        If blnNewPatient Then
            strSQLStatement = "exec sp_tbLabRequestMaster_select 'NEWPATIENT'"
        Else
            If Len(Trim(edtRequestID)) > 0 Then
                strSQLStatement = "exec sp_tbLabRequestMaster_select '" & edtPatientID & "','" & edtRequestID & "'"
            Else
                strSQLStatement = "exec sp_tbLabRequestMaster_select NULL,NULL"
            End If
        End If
        
        Set rsedtRequestReq = myConnection.Execute(strSQLStatement)

        numRowDemo = .Rows

        .Rows = numRowDemo + rsedtRequestReq.Fields.Count

        'set design and label of grid
        For intRowCtrReq = 0 To rsedtRequestReq.Fields.Count - 1
            .Row = numRowDemo + intRowCtrReq
'            .Col = 0
'            .CellAlignment = 0
            
            .Col = 1
            .CellAlignment = 1

            .TextMatrix(.Row, 2) = rsedtRequestReq.Fields(intRowCtrReq).Name
            .TextMatrix(.Row, 0) = FindSetting(.TextMatrix(.Row, 2), .TextMatrix(.Row, 2))
            .TextMatrix(.Row, 3) = FindSetting("ENA_" & .TextMatrix(.Row, 2), "0")
            .TextMatrix(.Row, 4) = FindSetting("VIS_" & .TextMatrix(.Row, 2), "1")
            .TextMatrix(.Row, 5) = FindSetting("MAN_" & .TextMatrix(.Row, 2), "1")
            .TextMatrix(.Row, 6) = FindSetting("TYPE_" & .TextMatrix(.Row, 2), "TEXT")
            .TextMatrix(.Row, 7) = IIf((RTrim(.TextMatrix(.Row, 6)) <> "TEXT"), FindSetting("SEL_" & .TextMatrix(.Row, 2), ""), "")
            
            If blnWithSettingsReq = False And .TextMatrix(.Row, 3) = "1" Then
                blnWithSettingsReq = True
            End If
        
            If Len(START_CUR) > 0 Then
                If .TextMatrix(.Row, 0) = START_CUR And .TextMatrix(.Row, 3) = "1" And .TextMatrix(.Row, 4) = "1" Then
                    .TextMatrix(.Row, 9) = "1"
                End If
            End If
    
            .CellBackColor = IIf(.TextMatrix(.Row, 3) = "1", &H80000005, &H80000011)
        
            If Not blnNewPatient Then
                If rsedtRequestReq.RecordCount > 0 Then
                    .TextMatrix(.Row, 1) = IIf(IsNull(rsedtRequestReq.Fields(intRowCtrReq).Value), "", Trim(rsedtRequestReq.Fields(intRowCtrReq).Value))
                    If intRowCtrReq = 20 Then
                        txtRemarks.Text = IIf(IsNull(rsedtRequestReq.Fields(intRowCtrReq).Value), "", Trim(rsedtRequestReq.Fields(intRowCtrReq).Value))
                    End If
                    .Col = 11
                    .CellAlignment = 1
                    .TextMatrix(.Row, 11) = IIf(IsNull(rsedtRequestReq.Fields(intRowCtrReq).Value), "", Trim(rsedtRequestReq.Fields(intRowCtrReq).Value))
                Else
                    .TextMatrix(.Row, 11) = "NEWREQUEST"
                End If
            Else
                If Trim(.TextMatrix(.Row, 2)) = "ReqBy" Then
                    .TextMatrix(.Row, 1) = edtUserName & "|" & edtUserID
                    .TextMatrix(.Row, 11) = "NEWPATIENT"
                Else
                    .TextMatrix(.Row, 11) = "NEWPATIENT"
                End If
            End If
                    
            
            If InStrRev(.TextMatrix(.Row, 1), "|") > 0 Then
                .TextMatrix(.Row, 10) = Mid(.TextMatrix(.Row, 1), InStrRev(.TextMatrix(.Row, 1), "|") + 1)
                .TextMatrix(.Row, 1) = Mid(.TextMatrix(.Row, 1), 1, InStrRev(.TextMatrix(.Row, 1), "|") - 1)
            End If
          
            
            If AUTO_COMP_AGE = "1" Then
                If .TextMatrix(.Row, 2) = "Age" Then
                    intRowReqAge = .Row
                    If Len(edtRequestID) > 0 Then
                        If Len(Trim(.TextMatrix(.Row, 1))) = 0 And Len(.TextMatrix(5, 1)) > 0 And IsDate(.TextMatrix(5, 1)) Then
                            .TextMatrix(.Row, 1) = CompAge(rsedtRequestReq.Fields("Date").Value, .TextMatrix(5, 1))
                        End If
                    Else
                        If Len(Trim(.TextMatrix(.Row, 1))) = 0 And Len(.TextMatrix(5, 1)) > 0 And IsDate(.TextMatrix(5, 1)) Then
                            .TextMatrix(.Row, 1) = CompAge(Date, .TextMatrix(5, 1))
                        End If
                    End If
                End If
            End If
                          
            .RowHeight(.Row) = IIf(.TextMatrix(.Row, 4) = "1", 350, 0)
            
            If .TextMatrix(.Row, 0) = strStartRowEdtReq Then
                .TextMatrix(.Row, 9) = "1"
            End If
    
        Next
        
        If Not blnNewPatient And rsedtRequestReq.RecordCount = 0 And Len(Trim(edtRequestID)) = 0 Then
            If Len(Trim(.TextMatrix(5, 1))) > 0 And IsDate(.TextMatrix(5, 1)) Then
                If AUTO_COMP_AGE = "1" Then
                    .TextMatrix(intRowReqAge, 1) = CompAge(Date, .TextMatrix(5, 1))
                End If
            End If
        End If
        
        .Col = 1
    End With
End Sub

Public Property Let setPatientID(ByVal vNewValue As String)
    edtPatientID = Trim(vNewValue)
End Property

Public Property Let setPatientHospID(ByVal vNewValue As String)
    edtPatientHospID = Trim(vNewValue)
End Property

Public Property Let SetRequestID(ByVal vNewValue As String)
    edtRequestID = Trim(vNewValue)
End Property

Public Property Let setSectionID(ByVal vNewValue As String)
    edtSectionID = Trim(vNewValue)
End Property

Public Property Let setSpecimen(ByVal vNewValue As String)
    edtSpecimen = Trim(vNewValue)
End Property

Public Property Let setSampleNumber(ByVal vNewValue As String)
    edtSampleNumber = Trim(vNewValue)
End Property

Public Property Let setUserID(ByVal vNewValue As String)
    edtUserID = Trim(vNewValue)
End Property

Public Property Let setUserName(ByVal vNewValue As String)
    edtUserName = Trim(vNewValue)
End Property

Public Property Let setPatientName(ByVal vNewValue As String)
    newPatientName = vNewValue
End Property

Public Property Let setBlnNewPatient(ByVal vNewValue As Boolean)
    blnNewPatient = vNewValue
End Property

Private Sub Form_Unload(Cancel As Integer)
            
    If blnSaveButtonClick = True Then
        If LTrim(lblLocked.Caption) <> "Locked. Sample is already opened by another user" Then
            ReqUnlock
        End If
        If Len(Trim(edtPatientHospID)) > 0 Then
            delTempPatientDemo
        End If
        blnSaveButtonClick = False
        blnNewPatient = False
        edtPatientID = ""
        edtPatientHospID = ""
        edtRequestID = ""
        edtSectionID = ""
        edtSpecimen = ""
        edtSampleNumber = ""
        edtUserID = ""
        edtUserName = ""
        newPatientName = ""
        currSampleStatus = ""
    Else
        If Not isExiting Then
            If Not blnFormIsLocked Then
                If MsgBox("Are you sure you want to exit without saving?", vbYesNo + vbQuestion, "Request") = vbYes Then
                    If LTrim(lblLocked.Caption) <> "Locked. Sample is already opened by another user" Then
                        ReqUnlock
                    End If
                    If Len(Trim(edtPatientHospID)) > 0 Then
                        delTempPatientDemo
                    End If
                    blnNewPatient = False
                    edtPatientID = ""
                    edtPatientHospID = ""
                    edtRequestID = ""
                    edtSectionID = ""
                    edtSpecimen = ""
                    edtSampleNumber = ""
                    edtUserID = ""
                    edtUserName = ""
                    newPatientName = ""
                    currSampleStatus = ""
                Else
                    tlbarReq.Enabled = True
                    If txtmsfgExam.Visible = True And txtmsfgExam.Enabled = True Then
                        txtmsfgExam.SetFocus
                    Else
                        txtmsfgResult.SetFocus
                    End If
                    Cancel = 1
                End If
            Else
                ReqUnlock
                blnNewPatient = False
                edtPatientID = ""
                edtPatientHospID = ""
                edtRequestID = ""
                edtSectionID = ""
                edtSpecimen = ""
                edtSampleNumber = ""
                edtUserID = ""
                edtUserName = ""
                newPatientName = ""
                currSampleStatus = ""
            End If
        End If
    End If
End Sub


Private Sub mnucmdClose_Click()
    Unload Me
End Sub

Private Sub mnucmdCnlReq_Click()
'    If MsgBox("Are you sure you want to cancel this request?", vbInformation + vbYesNo, "LMS") = vbYes Then
'
'
'    End If
End Sub

Private Sub mnucmdPrntLbl_Click()
    tlbarReq.Enabled = False
    mnucmdSave.Enabled = False
    mnucmdPrntLbl.Enabled = False
    mnucmdRmvTest.Enabled = False
    mnucmdClose.Enabled = False
    With msfgExam
        If Len(.TextMatrix(.Row, 8)) > 0 Then
            modClassMain.printBarCode .TextMatrix(.Row, 8)
            msfgExam_EnterCell
        End If
    End With
    tlbarReq.Enabled = True
    mnucmdSave.Enabled = True
    mnucmdPrntLbl.Enabled = True
    mnucmdRmvTest.Enabled = True
    mnucmdClose.Enabled = True
End Sub

Private Sub mnucmdPrntReqFrm_Click()
    If Len(Trim(edtRequestID)) > 0 And Len(Trim(edtSectionID)) > 0 Then
        OpenMainReport App.Path & "\Reports\MSSQL SERVER\Request.rpt", edtRequestID, edtUserID, edtSectionID, edtSampleNumber
    End If
End Sub

Private Sub mnucmdRmvTest_Click()
    Dim rsDelTest As New ADODB.Recordset
    Dim intColClear As Integer
    tlbarReq.Enabled = False
    
    mnucmdSave.Enabled = False
    mnucmdPrntLbl.Enabled = False
    mnucmdRmvTest.Enabled = False
    mnucmdClose.Enabled = False

    With msfgExam
        If Len(.TextMatrix(.Row, 9)) > 0 Then
            msfgExamDel.TextMatrix(msfgExamDel.Rows - 1, 0) = msfgResult.TextMatrix(15, 1)     'RequestID
            msfgExamDel.TextMatrix(msfgExamDel.Rows - 1, 1) = .TextMatrix(.Row, 7)             'SectionID
            msfgExamDel.TextMatrix(msfgExamDel.Rows - 1, 2) = .TextMatrix(.Row, 8)             'SampleID
            msfgExamDel.TextMatrix(msfgExamDel.Rows - 1, 3) = .TextMatrix(.Row, 1)             'TestID
            msfgExamDel.TextMatrix(msfgExamDel.Rows - 1, 4) = .TextMatrix(.Row, 2)             'Exam Name
            msfgExamDel.TextMatrix(msfgExamDel.Rows - 1, 5) = .TextMatrix(.Row, 6)             'Status
            msfgExamDel.TextMatrix(msfgExamDel.Rows - 1, 6) = .TextMatrix(.Row, 9)             'RVID
            msfgExamDel.TextMatrix(msfgExamDel.Rows - 1, 7) = .TextMatrix(.Row, 11)            'Priority
            msfgExamDel.TextMatrix(msfgExamDel.Rows - 1, 8) = .TextMatrix(.Row, 12)            'PanelID
            
            If Len(currSampleStatus) = 0 Then
                currSampleStatus = .TextMatrix(.Row, 6)
            End If
            
            msfgExamDel.Rows = msfgExamDel.Rows + 1
            DelExamDet .TextMatrix(.Row, 7), .TextMatrix(.Row, 1)
            .RemoveItem (.Row)
        Else
            DelExamDet .TextMatrix(.Row, 7), .TextMatrix(.Row, 1)
            .RemoveItem (.Row)
        End If
        
'
'        If .Rows > 3 Then
'            If Len(.TextMatrix(.Row, 8)) > 0 Then
'                If .TextMatrix(.Row, 6) < 2 Then
'                    If MsgBox("Are you sure you want to remove this item?", vbQuestion + vbYesNo, "LMS") = vbYes Then
'                        strSQLStatement = "EXEC sp_DelValuesByTest_delete '" & msfgResult.TextMatrix(15, 1) & "'," & _
'                                                "'" & .TextMatrix(.Row, 7) & "','" & .TextMatrix(.Row, 8) & "'," & _
'                                                "'" & .TextMatrix(.Row, 1) & "','" & edtUserID & "'"
'                        myConnection.Execute (strSQLStatement)
'
'                        DelExamDet .TextMatrix(.Row, 7), .TextMatrix(.Row, 1)
'                        .RemoveItem (.Row)
'                    End If
'                Else
'                    MsgBox "Unable to delete this test, sample already processed", vbOKOnly + vbInformation, "LMS"
'                End If
'            Else
'                If Len(msfgExam.TextMatrix(msfgExam.Row, 1)) > 0 Then
'                    DelExamDet msfgExam.TextMatrix(msfgExam.Row, 7), msfgExam.TextMatrix(msfgExam.Row, 1)
'                End If
'                If msfgExam.Rows > 2 Then
'                    If Len(msfgExam.TextMatrix(msfgExam.Row, 2)) > 0 Then
'                        msfgExam.RemoveItem (msfgExam.Row)
'                        msfgExam.Row = msfgExam.Rows - 1
'                        msfgExam_EnterCell
'                    End If
'                Else
'                    For intColClear = 0 To msfgExam.Cols - 1
'                        msfgExam.TextMatrix(msfgExam.Row, intColClear) = ""
'                    Next
'                End If
'            End If
'        Else
'            If Len(.TextMatrix(.Row, 8)) > 0 Then
'                MsgBox "Unable to remove all test request", vbInformation + vbOKOnly, "LMS"
'            Else
'                If Len(msfgExam.TextMatrix(msfgExam.Row, 1)) > 0 Then
'                    DelExamDet msfgExam.TextMatrix(msfgExam.Row, 7), msfgExam.TextMatrix(msfgExam.Row, 1)
'                End If
'                If msfgExam.Rows > 2 Then
'                    If Len(msfgExam.TextMatrix(msfgExam.Row, 2)) > 0 Then
'                        msfgExam.RemoveItem (msfgExam.Row)
'                        msfgExam.Row = msfgExam.Rows - 1
'                        msfgExam_EnterCell
'                    End If
'                Else
'                    For intColClear = 0 To msfgExam.Cols - 1
'                        msfgExam.TextMatrix(msfgExam.Row, intColClear) = ""
'                    Next
'                End If
'            End If
'        End If
    End With

'    msfgExam_LeaveCell
'    msfgExam.SetFocus
    msfgExam_EnterCell
    msfgExam.ColSel = msfgExam.Cols - 1
                
    tlbarReq.Enabled = True
    mnucmdSave.Enabled = True
    mnucmdPrntLbl.Enabled = True
    mnucmdRmvTest.Enabled = True
    mnucmdClose.Enabled = True
End Sub

Private Sub mnucmdSave_Click()
    Dim intRowCtrTlbar As Integer
    Dim blnNewOtherExam As Boolean
    Dim blnChangePriority As Boolean
    Dim blnChangeDemo As Boolean
    Dim blnChangeReq As Boolean
    Dim intRowChange As Integer
    Dim rsSave As New ADODB.Recordset
    
    Dim withSTAT As String
    
    Dim ctrmsfgExamDel As Integer
    Dim TestToRemove As String

    tlbarReq.Enabled = False
    mnucmdSave.Enabled = False
    mnucmdPrntLbl.Enabled = False
    mnucmdRmvTest.Enabled = False
    mnucmdClose.Enabled = False
    msfgResult_LeaveCell
    
    If Len(edtRequestID) = 0 Then
        If msfgExam.Rows = 2 Then
            MsgBox "No Test Requested", vbInformation + vbOKOnly, "LMS"
            If Not txtmsfgResult.Visible And Not cmbmsfgResult.Visible Then
                msfgExam_EnterCell
            Else
                msfgResult_EnterCell
            End If
            tlbarReq.Enabled = True
            Exit Sub
        End If
    End If
    
    If blnMandatoryClear Then
        If blnNewPatient Then
            With msfgResult
                strSQLStatement = "EXEC sp_PatientDemo_insert " & _
                    IIf((Len(Trim(.TextMatrix(1, 1))) > 0), "'" & .TextMatrix(1, 1) & "',", "NULL,") & _
                    IIf((Len(Trim(.TextMatrix(2, 1))) > 0), "'" & .TextMatrix(2, 1) & "',", "NULL,") & _
                    IIf((Len(Trim(.TextMatrix(3, 1))) > 0), "'" & .TextMatrix(3, 1) & "',", "NULL,") & _
                    IIf((Len(Trim(.TextMatrix(4, 1))) > 0), "'" & .TextMatrix(4, 1) & "',", "NULL,") & _
                    IIf((Len(Trim(.TextMatrix(5, 1))) > 0), "'" & .TextMatrix(5, 1) & "',", "NULL,") & _
                    IIf((Len(Trim(.TextMatrix(6, 1))) > 0), "'" & .TextMatrix(6, 1) & "',", "NULL,") & _
                    IIf((Len(Trim(.TextMatrix(7, 1))) > 0), "'" & .TextMatrix(7, 1) & "',", "NULL,") & _
                    IIf((Len(Trim(.TextMatrix(8, 1))) > 0), "'" & .TextMatrix(8, 1) & "',", "NULL,") & _
                    IIf((Len(Trim(.TextMatrix(9, 1))) > 0), "'" & .TextMatrix(9, 1) & "',", "NULL,") & _
                    IIf((Len(Trim(.TextMatrix(10, 1))) > 0), "'" & .TextMatrix(10, 1) & "',", "NULL,") & _
                    IIf((Len(Trim(.TextMatrix(11, 1))) > 0), "'" & .TextMatrix(11, 1) & "',", "NULL,") & _
                    IIf((Len(Trim(.TextMatrix(12, 1))) > 0), "'" & .TextMatrix(12, 1) & "',", "NULL,") & _
                    "'" & edtUserID & "'"
                                                                    
                'Debug.Print strSQLStatement
                Set rsSave = myConnection.Execute(strSQLStatement)
                
                If rsSave.RecordCount > 0 Then
                    .TextMatrix(0, 1) = Trim(rsSave.Fields("newPID").Value)
                End If
            
                withSTAT = CheckWithSTAT
                
                strSQLStatement = "EXEC sp_tbLabRequestMaster_insert " & _
                      IIf((Len(Trim(.TextMatrix(13, 1))) > 0), "'" & .TextMatrix(13, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(0, 1))) > 0), "'" & .TextMatrix(0, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(16, 1))) > 0), "'" & .TextMatrix(16, 1) & "',", "0,") & _
                            "'" & edtUserID & "'," & _
                      IIf((Len(Trim(.TextMatrix(19, 1))) > 0), "'" & .TextMatrix(19, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(20, 1))) > 0), "'" & IIf(.TextMatrix(20, 6) <> "SELECTION", _
                            .TextMatrix(20, 1), .TextMatrix(20, 10)) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(21, 1))) > 0), "'" & IIf(.TextMatrix(21, 6) <> "SELECTION", _
                            .TextMatrix(21, 1), .TextMatrix(21, 10)) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(22, 1))) > 0), "'" & .TextMatrix(22, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(23, 1))) > 0), "'" & .TextMatrix(23, 10) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(24, 1))) > 0), "'" & .TextMatrix(24, 1) & "',", "NULL,") & _
                            "'" & withSTAT & "'," & _
                      IIf((Len(Trim(.TextMatrix(26, 1))) > 0), "'" & .TextMatrix(26, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(27, 1))) > 0), "'" & .TextMatrix(27, 10) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(28, 1))) > 0), "'" & .TextMatrix(28, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(29, 1))) > 0), "'" & .TextMatrix(29, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(30, 1))) > 0), "'" & .TextMatrix(30, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(31, 1))) > 0), "'" & .TextMatrix(31, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(32, 1))) > 0), "'" & .TextMatrix(32, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(txtRemarks.Text)) > 0), "'" & txtRemarks.Text & "',", "'',") & _
                      "'" & intWorkstation & "'"
                    
                'Debug.Print strSQLStatement
                Set rsSave = myConnection.Execute(strSQLStatement)
                
                If rsSave.RecordCount > 0 Then
                    .TextMatrix(14, 1) = .TextMatrix(0, 1)
                    .TextMatrix(15, 1) = rsSave.Fields("RequestID").Value
                    .TextMatrix(17, 1) = Format(rsSave.Fields("RequestDate").Value, "MM/dd/yyyy Hh:mm")
                End If
            End With
        Else
            With msfgResult
                If Len(Trim(edtPatientHospID)) > 0 Then
                    blnChangeDemo = True
                Else
                    For intRowChange = 0 To 12
                        If RTrim(.TextMatrix(intRowChange, 1)) <> RTrim(.TextMatrix(intRowChange, 11)) Then
                            blnChangeDemo = True
                            Exit For
                        End If
                    Next
                End If
                
                For intRowChange = 13 To .Rows - 1
                    If InStr(1, .TextMatrix(intRowChange, 11), "|") > 0 Then
                        If RTrim(.TextMatrix(intRowChange, 10)) <> RTrim(Mid(.TextMatrix(intRowChange, 11), InStr(1, .TextMatrix(intRowChange, 11), "|") + 1)) Then
                            blnChangeReq = True
                            Exit For
                        End If
                    Else
                        If intRowChange = 33 Then
                            If RTrim(.TextMatrix(intRowChange, 11)) <> RTrim(txtRemarks.Text) Then
                                blnChangeReq = True
                                Exit For
                            End If
                        Else
                            If RTrim(.TextMatrix(intRowChange, 1)) <> RTrim(.TextMatrix(intRowChange, 11)) Then
                                blnChangeReq = True
                                Exit For
                            End If
                        End If
                    End If
                Next
            End With
            If blnChangeDemo Then
                With msfgResult
                    strSQLStatement = "EXEC sp_PatientDemo_update " & _
                        "'" & Trim(.TextMatrix(0, 1)) & "'," & _
                        IIf((Len(Trim(.TextMatrix(1, 1))) > 0), "'" & .TextMatrix(1, 1) & "',", "NULL,") & _
                        IIf((Len(Trim(.TextMatrix(2, 1))) > 0), "'" & .TextMatrix(2, 1) & "',", "NULL,") & _
                        IIf((Len(Trim(.TextMatrix(3, 1))) > 0), "'" & .TextMatrix(3, 1) & "',", "NULL,") & _
                        IIf((Len(Trim(.TextMatrix(4, 1))) > 0), "'" & .TextMatrix(4, 1) & "',", "NULL,") & _
                        IIf((Len(Trim(.TextMatrix(5, 1))) > 0), "'" & .TextMatrix(5, 1) & "',", "NULL,") & _
                        IIf((Len(Trim(.TextMatrix(6, 1))) > 0), "'" & .TextMatrix(6, 1) & "',", "NULL,") & _
                        IIf((Len(Trim(.TextMatrix(7, 1))) > 0), "'" & .TextMatrix(7, 1) & "',", "NULL,") & _
                        IIf((Len(Trim(.TextMatrix(8, 1))) > 0), "'" & .TextMatrix(8, 1) & "',", "NULL,") & _
                        IIf((Len(Trim(.TextMatrix(9, 1))) > 0), "'" & .TextMatrix(9, 1) & "',", "NULL,") & _
                        IIf((Len(Trim(.TextMatrix(10, 1))) > 0), "'" & .TextMatrix(10, 1) & "',", "NULL,") & _
                        IIf((Len(Trim(.TextMatrix(11, 1))) > 0), "'" & .TextMatrix(11, 1) & "',", "NULL,") & _
                        IIf((Len(Trim(.TextMatrix(12, 1))) > 0), "'" & .TextMatrix(12, 1) & "',", "NULL,") & _
                        "'" & edtUserID & "'"
                End With
                
                Set rsSave = myConnection.Execute(strSQLStatement)
            End If
            If Len(Trim(edtRequestID)) > 0 Then
                withSTAT = CheckWithSTAT
                If blnChangeReq Then
                    With msfgResult
                        strSQLStatement = "EXEC sp_tbLabRequestMaster_update " & _
                            IIf((Len(Trim(.TextMatrix(13, 1))) > 0), "'" & .TextMatrix(13, 1) & "',", "NULL,") & _
                            IIf((Len(Trim(.TextMatrix(0, 1))) > 0), "'" & .TextMatrix(0, 1) & "',", "NULL,") & _
                            IIf((Len(Trim(.TextMatrix(15, 1))) > 0), "'" & .TextMatrix(15, 1) & "',", "NULL,") & _
                            IIf((Len(Trim(.TextMatrix(16, 1))) > 0), "'" & .TextMatrix(16, 1) & "',", "NULL,") & _
                                    "'" & edtUserID & "'," & _
                            IIf((Len(Trim(.TextMatrix(19, 1))) > 0), "'" & .TextMatrix(19, 1) & "',", "NULL,") & _
                            IIf((Len(Trim(.TextMatrix(20, 1))) > 0), "'" & IIf(.TextMatrix(20, 6) <> "SELECTION", _
                                    .TextMatrix(20, 1), .TextMatrix(20, 10)) & "',", "NULL,") & _
                            IIf((Len(Trim(.TextMatrix(21, 1))) > 0), "'" & IIf(.TextMatrix(21, 6) <> "SELECTION", _
                                    .TextMatrix(21, 1), .TextMatrix(21, 10)) & "',", "NULL,") & _
                            IIf((Len(Trim(.TextMatrix(22, 1))) > 0), "'" & .TextMatrix(22, 1) & "',", "NULL,") & _
                            IIf((Len(Trim(.TextMatrix(23, 1))) > 0), "'" & .TextMatrix(23, 10) & "',", "NULL,") & _
                            IIf((Len(Trim(.TextMatrix(24, 1))) > 0), "'" & .TextMatrix(24, 1) & "',", "NULL,") & _
                                    "'" & withSTAT & "'," & _
                            IIf((Len(Trim(.TextMatrix(26, 1))) > 0), "'" & .TextMatrix(26, 1) & "',", "NULL,") & _
                            IIf((Len(Trim(.TextMatrix(27, 1))) > 0), "'" & .TextMatrix(27, 10) & "',", "NULL,") & _
                            IIf((Len(Trim(.TextMatrix(28, 1))) > 0), "'" & .TextMatrix(28, 1) & "',", "NULL,") & _
                            IIf((Len(Trim(.TextMatrix(29, 1))) > 0), "'" & .TextMatrix(29, 1) & "',", "NULL,") & _
                            IIf((Len(Trim(.TextMatrix(30, 1))) > 0), "'" & .TextMatrix(30, 1) & "',", "NULL,") & _
                            IIf((Len(Trim(.TextMatrix(31, 1))) > 0), "'" & .TextMatrix(31, 1) & "',", "NULL,") & _
                            IIf((Len(Trim(.TextMatrix(32, 1))) > 0), "'" & .TextMatrix(32, 1) & "',", "NULL,") & _
                            IIf((Len(Trim(txtRemarks.Text)) > 0), "'" & txtRemarks.Text & "'", "''")
                                                
                        Set rsSave = myConnection.Execute(strSQLStatement)
                        
                    End With
                End If
            Else
                withSTAT = CheckWithSTAT
                
                With msfgResult
                    strSQLStatement = "EXEC sp_tbLabRequestMaster_insert " & _
                      IIf((Len(Trim(.TextMatrix(13, 1))) > 0), "'" & .TextMatrix(13, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(0, 1))) > 0), "'" & .TextMatrix(0, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(16, 1))) > 0), "'" & .TextMatrix(16, 1) & "',", "0,") & _
                            "'" & edtUserID & "'," & _
                      IIf((Len(Trim(.TextMatrix(19, 1))) > 0), "'" & .TextMatrix(19, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(20, 1))) > 0), "'" & IIf(.TextMatrix(20, 6) <> "SELECTION", _
                            .TextMatrix(20, 1), .TextMatrix(20, 10)) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(21, 1))) > 0), "'" & IIf(.TextMatrix(21, 6) <> "SELECTION", _
                            .TextMatrix(21, 1), .TextMatrix(21, 10)) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(22, 1))) > 0), "'" & .TextMatrix(22, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(23, 1))) > 0), "'" & .TextMatrix(23, 10) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(24, 1))) > 0), "'" & .TextMatrix(24, 1) & "',", "NULL,") & _
                            "'" & withSTAT & "'," & _
                      IIf((Len(Trim(.TextMatrix(26, 1))) > 0), "'" & .TextMatrix(26, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(27, 1))) > 0), "'" & .TextMatrix(27, 10) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(28, 1))) > 0), "'" & .TextMatrix(28, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(29, 1))) > 0), "'" & .TextMatrix(29, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(30, 1))) > 0), "'" & .TextMatrix(30, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(31, 1))) > 0), "'" & .TextMatrix(31, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(.TextMatrix(32, 1))) > 0), "'" & .TextMatrix(32, 1) & "',", "NULL,") & _
                      IIf((Len(Trim(txtRemarks.Text)) > 0), "'" & txtRemarks.Text & "',", "'',") & _
                      "'" & intWorkstation & "'"
                                            
                    Set rsSave = myConnection.Execute(strSQLStatement)
                                            
                    If rsSave.RecordCount > 0 Then
                        .TextMatrix(14, 1) = .TextMatrix(0, 1)
                        .TextMatrix(15, 1) = rsSave.Fields("RequestID").Value
                        .TextMatrix(17, 1) = Format(rsSave.Fields("RequestDate").Value, "MM/dd/yyyy Hh:mm")
                    End If
                    
                End With
            End If
        End If
        With msfgExam
            For intRowChange = 1 To .Rows - 1
                If Len(Trim(.TextMatrix(intRowChange, 6))) > 0 Then
                    If Len(Trim(edtRequestID)) > 0 And Len(Trim(edtSampleNumber)) > 0 Then
                        If Trim(.TextMatrix(intRowChange, 7)) = edtSectionID Then
                            .TextMatrix(intRowChange, 8) = Trim(edtSampleNumber)
                            .TextMatrix(intRowChange, 0) = .TextMatrix(1, 6)
                            .TextMatrix(intRowChange, 6) = .TextMatrix(1, 6)
                        Else
                            If Not blnNewOtherExam Then
                                blnNewOtherExam = True
                            End If
                        End If
                    End If
                    If Len(.TextMatrix(intRowChange, 9)) = 0 Then
                        strSQLStatement = "EXEC sp_tbLabRequestValues_insert " & _
                                        "'" & msfgResult.TextMatrix(15, 1) & "'," & _
                                        "'" & .TextMatrix(intRowChange, 7) & "'," & _
                                        "'" & .TextMatrix(intRowChange, 1) & "'," & _
                                        "'" & IIf(Len(currSampleStatus) = 0, .TextMatrix(intRowChange, 6), currSampleStatus) & "'," & _
                                        "'" & .TextMatrix(intRowChange, 8) & "'," & _
                                        "'" & edtUserID & "'," & _
                                        "'" & IIf(.TextMatrix(intRowChange, 11) = Chr(254), "1", "0") & "'," & _
                                        "''"
                        myConnection.Execute (strSQLStatement)
                    End If
                End If
            Next
                        
            If .Rows > 2 Then
                If Len(Trim(edtSampleNumber)) = 0 Then
                    If AUTO_CHECK_IN = "1" Then
                        strSQLStatement = "EXEC sp_CheckInByReq_update '" & msfgResult.TextMatrix(15, 1) & "'," & _
                                            "'" & Format(msfgResult.TextMatrix(17, 1), "MM/dd/yyyy") & "'," & edtUserID

                        myConnection.Execute (strSQLStatement)
                        
                        If AUTO_PRINT_BC = "1" Then
                            printBarCode "NEWPATIENT", msfgResult.TextMatrix(15, 1), LABEL_NUM_COPY
                        End If
                        If AUTOPRNTREQFRM = "1" Then
                            OpenMainReport App.Path & "\Reports\MSSQL SERVER\Request.rpt", msfgResult.TextMatrix(15, 1), edtUserID, "0", ""
                        End If
                        If DISPTESTREQSUM = "1" Then
                            showSummaryRequest msfgResult.TextMatrix(15, 1)
                        End If
                    End If
                Else
                    If blnNewOtherExam Then
                        If AUTO_CHECK_IN = "1" Then
                            strSQLStatement = "EXEC sp_CheckInByReq_update '" & msfgResult.TextMatrix(15, 1) & "'," & _
                                            "'" & Format(msfgResult.TextMatrix(17, 1), "MM/dd/yyyy") & "'," & edtUserID
                            myConnection.Execute (strSQLStatement)
                        End If
                        blnNewOtherExam = False
                    End If
                End If
                
                If Len(Trim(edtSampleNumber)) > 0 Then
                    strSQLStatement = ""
                    For intRowChange = 1 To .Rows - 1
                        If .TextMatrix(intRowChange, 11) <> .TextMatrix(intRowChange, 13) Then
                            If Len(strSQLStatement) = 0 Then
                                strSQLStatement = "EXEC sp_tblabRequestValues_backup '" & edtRequestID & "','" & _
                                        edtSampleNumber & "','" & edtSectionID & "','" & edtUserID & "'" & vbCrLf
                                strSQLStatement = strSQLStatement & "EXEC sp_tbLabRequestValues_update '" & edtRequestID & "','" & _
                                        edtSampleNumber & "','" & .TextMatrix(intRowChange, 1) & "','" & _
                                        edtSectionID & "','" & IIf(.TextMatrix(intRowChange, 11) = Chr(254), "1", "0") & "','" & _
                                        edtUserID & "'" & vbCrLf
                            Else
                                strSQLStatement = strSQLStatement & "EXEC sp_tbLabRequestValues_update '" & edtRequestID & "','" & _
                                        edtSampleNumber & "','" & .TextMatrix(intRowChange, 1) & "','" & _
                                        edtSectionID & "','" & IIf(.TextMatrix(intRowChange, 11) = Chr(254), "1", "0") & "','" & _
                                        edtUserID & "'" & vbCrLf
                            End If
                        End If
                    Next
                    If Len(strSQLStatement) > 0 Then
                        myConnection.Execute (strSQLStatement)
                    End If
                End If
            End If
        
            If msfgExamDel.Rows > 2 Then
                For ctrmsfgExamDel = 1 To msfgExamDel.Rows - 1
                    If Len(msfgExamDel.TextMatrix(ctrmsfgExamDel, 4)) > 0 Then
                        TestToRemove = TestToRemove & vbCr & msfgExamDel.TextMatrix(ctrmsfgExamDel, 4)
                    End If
                Next
                If Len(TestToRemove) Then
                    TestToRemove = "You are going to remove Exam/s" & vbCr & TestToRemove & vbCr & vbCr & "Do You want to Proceed?"
                    If .Rows = 2 And Len(.TextMatrix(.Rows - 1, 1)) = 0 Then
                        If MsgBox("By Removing all Exam Requested, This Sample ID will be automatically tagged as  Cancelled, Do You want to Proceed?", _
                                    vbYesNo + vbInformation, "LMS") = vbYes Then
                                    
                            If Len(Trim(txtRemarks.Text)) > 0 Then
                                DelExam True
                                rPatientID = msfgResult.TextMatrix(0, 1)
                                rRequestID = msfgResult.TextMatrix(15, 1)
                                
                                blnSaveButtonClick = True
                                Unload Me
                                Unload frmSearchRequest
                            Else
                                MsgBox "Please provide a cancellation note in the remarks field", vbInformation, "LMS"
                                mnucmdSave.Enabled = True
                                mnucmdPrntLbl.Enabled = True
                                mnucmdRmvTest.Enabled = True
                                mnucmdClose.Enabled = True
                                tlbarReq.Enabled = True
                            End If
                        Else
                            mnucmdSave.Enabled = True
                            mnucmdPrntLbl.Enabled = True
                            mnucmdRmvTest.Enabled = True
                            mnucmdClose.Enabled = True
                            tlbarReq.Enabled = True
                        End If
                    Else
                        If MsgBox(TestToRemove, vbYesNo + vbInformation, "LMS") = vbYes Then
                            DelExam False
                            rPatientID = msfgResult.TextMatrix(0, 1)
                            rRequestID = msfgResult.TextMatrix(15, 1)
                            
                            blnSaveButtonClick = True
                            Unload Me
                            Unload frmSearchRequest
                        Else
                            mnucmdSave.Enabled = True
                            mnucmdPrntLbl.Enabled = True
                            mnucmdRmvTest.Enabled = True
                            mnucmdClose.Enabled = True
                            tlbarReq.Enabled = True
                        End If
                    End If
                End If
            Else
                rPatientID = msfgResult.TextMatrix(0, 1)
                rRequestID = msfgResult.TextMatrix(15, 1)
                
                blnSaveButtonClick = True
                Unload Me
                Unload frmSearchRequest
            End If
        End With
    Else
        msfgResult_EnterCell
        If Not txtmsfgResult.Visible And Not cmbmsfgResult.Visible Then
            msfgExam_EnterCell
        End If
        tlbarReq.Enabled = True
        mnucmdSave.Enabled = False
        mnucmdPrntLbl.Enabled = True
        mnucmdRmvTest.Enabled = True
        mnucmdClose.Enabled = True
    End If
    
End Sub

Private Sub DelExam(CancelAll As Boolean)
    Dim CancelReason As String
    Dim ctrmsfgExamDelRow As Integer
    
    '0 RequestID
    '1 SectionID
    '2 SampleID
    '3 TestID
    '4 Exam Name
    '5 Status
    '6 RVID
    '7 Priority
    '8 PanelID
    
    If CancelAll Then
        myConnection.Execute ("EXEC sp_RequestValues_Update '" & edtSampleNumber & "','7','" & edtUserID & "','','" & Replace(txtRemarks.Text, vbCrLf, " ") & "','0','0'")
    Else
        With msfgExamDel
            myConnection.Execute ("EXEC sp_tblabRequestValues_backup '" & edtRequestID & "','" & edtSampleNumber & "','" & edtSectionID & "','" & edtUserID & "'")
            For ctrmsfgExamDelRow = 1 To .Rows - 1
                If Len(.TextMatrix(ctrmsfgExamDelRow, 0)) > 0 Then
                    myConnection.Execute ("EXEC sp_RequestValuesCanceled_delete " & .TextMatrix(ctrmsfgExamDelRow, 6) & ",'" & edtSampleNumber & "','" & _
                                            .TextMatrix(ctrmsfgExamDelRow, 3) & "','" & .TextMatrix(ctrmsfgExamDelRow, 1) & "','" & edtUserID & "'")
                End If
            Next
        End With
    End If
    
End Sub

Private Sub msfgExam_Click()
    With msfgExam
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub msfgExam_EnterCell()
    msfgResult_LeaveCell
    With msfgExam
        .ColSel = .Cols - 1
        If .Col = 11 And (Len(.TextMatrix(.Row, 2)) > 0) Then
            If .TextMatrix(.Row, .Col) = Chr(111) Then
                .TextMatrix(.Row, .Col) = Chr(254)
            Else
                .TextMatrix(.Row, .Col) = Chr(111)
            End If

        End If
        
        .Col = 1

        If Len(Trim(.TextMatrix(.Row, 8))) > 2 Then
            tlbarReq.Buttons(2).Enabled = True
            mnucmdPrntLbl.Enabled = True
        Else
            tlbarReq.Buttons(2).Enabled = False
            mnucmdPrntLbl.Enabled = False
        End If
        
        If Len(.TextMatrix(.Row, 6)) > 0 Then
            If Val(.TextMatrix(.Row, 6)) < 2 Then
                mnucmdRmvTest.Enabled = True
                tlbarReq.Buttons(3).Enabled = True
            Else
                mnucmdRmvTest.Enabled = False
                tlbarReq.Buttons(3).Enabled = False
            End If
        Else
            mnucmdRmvTest.Enabled = False
            tlbarReq.Buttons(3).Enabled = False
        End If
    End With
End Sub

Private Sub msfgExam_GotFocus()
    msfgExam.Col = 1
    msfgExam.ColSel = msfgExam.Cols - 1
End Sub

Private Sub msfgExam_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intColClear As Integer

    If KeyCode = 46 Then
        If tlbarReq.Buttons(3).Enabled = True Then
            mnucmdRmvTest_Click
        End If
    End If
    msfgExam.SetFocus
    msfgExam.Col = 1
    msfgExam.ColSel = msfgExam.Cols - 1
End Sub

Private Sub msfgExam_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Then
        msfgResult_EnterCell
        KeyAscii = 0
    End If
End Sub


Private Sub msfgExam_LeaveCell()
    mnucmdRmvTest.Enabled = False
    tlbarReq.Buttons(3).Enabled = False
End Sub

Private Sub msfgResult_Click()
    msfgResult_EnterCell
End Sub

Private Sub msfgResult_EnterCell()
    Dim intRowFieldType As Integer
    msfgExam_LeaveCell
    With msfgResult
        If .Col = 1 Then
            txtmsfgResult.Visible = False
            txtmsfgResult.Enabled = False
            cmbmsfgResult.Visible = False
            cmbmsfgResult.Enabled = False
            If .TextMatrix(.Row, 3) = "1" And .TextMatrix(.Row, 4) = "1" Then
                .CellAlignment = 0
                If .TextMatrix(.Row, 6) = "TEXT" Or .TextMatrix(.Row, 6) = "SELECTION" Then
                    txtmsfgResult.Move (.CellLeft + .Left), (.CellTop + .Top), .CellWidth, .CellHeight
                    txtmsfgResult.Visible = True
                    txtmsfgResult.Enabled = True
                    txtmsfgResult.Text = .TextMatrix(.Row, .Col)
                    txtmsfgResult.MaxLength = setLenTxtBox(.TextMatrix(.Row, 2))
                    txtmsfgResult.SetFocus
                Else
                    cmbmsfgResult.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
                    cmbmsfgResult.Visible = True
                    cmbmsfgResult.Enabled = True
                    setListcmbmsfgReqDet .TextMatrix(.Row, 0)
                    If Len(.TextMatrix(.Row, .Col)) > 0 Then
                        For intRowFieldType = 0 To cmbmsfgResult.ListCount - 1
                            If .TextMatrix(.Row, 1) = cmbmsfgResult.List(intRowFieldType) Then
                                cmbmsfgResult.ListIndex = intRowFieldType
                            End If
                        Next
                    End If
                    cmbmsfgResult.SetFocus
                End If
            End If
        Else
            If txtmsfgResult.Visible = True And txtmsfgResult.Enabled = True Then
                txtmsfgResult.SetFocus
                highLight
            ElseIf cmbmsfgResult.Visible = True And cmbmsfgResult.Enabled = True Then
                cmbmsfgResult.SetFocus
            End If
        End If
    End With
End Sub
 

Private Sub msfgResult_LeaveCell()
    With msfgResult
        If .Col = 1 And .TextMatrix(.Row, 3) = "1" And .TextMatrix(.Row, 4) = "1" Then
            If .TextMatrix(.Row, 6) = "TEXT" Or .TextMatrix(.Row, 6) = "SELECTION" Then
                .TextMatrix(.Row, .Col) = txtmsfgResult.Text
                If Len(Trim(txtmsfgResult.Text)) = 0 Then .TextMatrix(.Row, 10) = ""
            Else
                .TextMatrix(.Row, .Col) = cmbmsfgResult.Text
            End If
            If .TextMatrix(.Row, 2) = "DOB" And Len(.TextMatrix(.Row, 1)) > 0 Then
                If Not IsDate(txtmsfgResult.Text) Then
                    .TextMatrix(.Row, 1) = ""
                    .TextMatrix(intRowReqAge, 1) = ""
                ElseIf Val(Year(.TextMatrix(.Row, 1))) > Val(Year(Now)) Then
                    .TextMatrix(.Row, 1) = ""
                    .TextMatrix(intRowReqAge, 1) = ""
                Else
                    .TextMatrix(.Row, 1) = Format(txtmsfgResult.Text, "MM/dd/yyyy")

                    If AUTO_COMP_AGE = "1" Then
                        .TextMatrix(intRowReqAge, 1) = CompAge(IIf(Len(.TextMatrix(17, 1)) = 0, Date, _
                            Format(.TextMatrix(17, 1), "mm/dd/yyyy")), txtmsfgResult.Text)
                        If Val(.TextMatrix(intRowReqAge, 1)) > 100 Then
                            If MsgBox("Age is greater than 100.", vbInformation + vbOKCancel, "LMS") = vbCancel Then
                                .TextMatrix(.Row, 1) = ""
                                .TextMatrix(intRowReqAge, 1) = ""
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        If txtmsfgResult.Visible = True Then
            txtmsfgResult.Visible = False
            txtmsfgResult.Enabled = False
        End If
        
        If cmbmsfgResult.Visible = True Then
            cmbmsfgResult.Visible = False
            cmbmsfgResult.Enabled = False
        End If
    End With
End Sub

Private Sub msfgResult_Scroll()
    txtmsfgResult.Visible = False
    txtmsfgResult.Enabled = False
    cmbmsfgResult.Visible = False
    cmbmsfgResult.Enabled = False
End Sub

Private Sub SaveDemoRequest()
    Dim blnChangeDemo As Boolean
    Dim blnChangeReq As Boolean
    Dim intRowChange As Integer
    Dim rsSave As New ADODB.Recordset
End Sub

Private Sub tlbarReq_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim intRowCtrTlbar As Integer
    Dim blnNewOtherExam As Boolean
    Select Case Button.Index
        Case 1
            mnucmdSave_Click
        Case 2
            mnucmdPrntLbl_Click
        Case 3
            mnucmdRmvTest_Click
        Case 4
            mnucmdPrntReqFrm_Click
        Case 5
            mnucmdClose_Click
    End Select
    
End Sub

Private Sub txtmsfgExam_GotFocus()
    highLight
    highLight_bg
    msfgResult_LeaveCell
    msfgExam_LeaveCell
    txtmsfgResult.Visible = False
End Sub

Private Sub txtmsfgExam_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
        cmdExamMenu_Click
        highLight
        highLight_bg
    End If

End Sub

Private Sub txtmsfgExam_KeyPress(KeyAscii As Integer)
    Dim strSelectionValue As String
    Dim rsSearchExam As New ADODB.Recordset
    Dim intSelValueColCtr As Integer
    Dim KP_arryExamSelected() As String
    Dim KP_arryExamDetFromRemoved() As String
    
    If KeyAscii = 13 Then
    
        If Len(Trim(txtmsfgExam.Text)) = 0 Then
            If Not blnNoExistingTest Then
                strSelectionValue = showfrmExamSelection(txtmsfgExam.Text, edtUserID, edtSpecimen)
            Else
                strSelectionValue = showfrmExamSelection(txtmsfgExam.Text, edtUserID, "")
            End If
        Else
            strSQLStatement = "EXEC sp_ReqExam_Select " & edtUserID & ",'" & txtmsfgExam.Text & "'"

            Set rsSearchExam = myConnection.Execute(strSQLStatement)
            
            If Not blnNewPatient And Not blnNoExistingTest Then
                If Len(edtSectionID) = 0 Then
                    rsSearchExam.Filter = "Specimen = '" & edtSpecimen & "'"
                Else
                    rsSearchExam.Filter = "Specimen = '" & edtSpecimen & "' AND SectionID = '" & edtSectionID & "'"
                End If
            End If
            
            If rsSearchExam.RecordCount > 0 Then
                If rsSearchExam.RecordCount > 1 Then
                    If Not blnNoExistingTest Then
                        strSelectionValue = showfrmExamSelection(txtmsfgExam.Text, edtUserID, edtSpecimen)
                    Else
                        strSelectionValue = showfrmExamSelection(txtmsfgExam.Text, edtUserID, "")
                    End If
                Else
                    For intSelValueColCtr = 0 To rsSearchExam.Fields.Count - 1
                        strSelectionValue = strSelectionValue & rsSearchExam.Fields(intSelValueColCtr) & "|"
                    Next
                    strSelectionValue = Mid(strSelectionValue, 1, Len(strSelectionValue) - 1)
                End If
            Else
                If Not blnNoExistingTest Then
                    strSelectionValue = showfrmExamSelection(txtmsfgExam.Text, edtUserID, edtSpecimen, edtSectionID)
                Else
                    strSelectionValue = showfrmExamSelection(txtmsfgExam.Text, edtUserID, "", "")
                End If
            End If
        End If
        
        If Len(strSelectionValue) > 0 Then
            KP_arryExamSelected = Split(strSelectionValue, "|")
            If Not blnExamOccured(strSelectionValue) Then
                With msfgExam
                    .Row = .Rows - 1
                    
                    .TextMatrix(.Row, 0) = "0"
                    .CellAlignment = 1
                    .TextMatrix(.Row, 1) = KP_arryExamSelected(0)
                    .TextMatrix(.Row, 2) = KP_arryExamSelected(1)
                    .TextMatrix(.Row, 3) = KP_arryExamSelected(2)
                    .TextMatrix(.Row, 4) = KP_arryExamSelected(5)
                    .TextMatrix(.Row, 5) = ""
                    .TextMatrix(.Row, 6) = "0"
                    .TextMatrix(.Row, 7) = KP_arryExamSelected(4)
                    .TextMatrix(.Row, 8) = ""
                    .TextMatrix(.Row, 9) = ""
                    .TextMatrix(.Row, 10) = KP_arryExamSelected(3)
                    .Col = 11
                    .CellFontName = "Wingdings"
                    .CellFontBold = True
                    .CellFontSize = 12
                    .CellAlignment = 4
                    .TextMatrix(.Row, 11) = IIf(chkStat.Value = 0, Chr(111), Chr(254))
                    .Col = 0
                    .CellForeColor = vbWhite
                
                    If Len(edtRequestID) > 0 Then
                        .TextMatrix(.Row, 8) = edtSampleNumber
                        KP_arryExamDetFromRemoved = Split(RemoveFromTestToDel(KP_arryExamSelected(0)), "|")
                        If UBound(KP_arryExamDetFromRemoved) > 0 Then
                            .TextMatrix(.Row, 0) = KP_arryExamDetFromRemoved(0)
                            .TextMatrix(.Row, 6) = KP_arryExamDetFromRemoved(0)
                            .TextMatrix(.Row, 9) = KP_arryExamDetFromRemoved(1)
                            .TextMatrix(.Row, 11) = KP_arryExamDetFromRemoved(2)
                            .TextMatrix(.Row, 12) = KP_arryExamDetFromRemoved(3)
                            .Col = 0
                            .CellPictureAlignment = 4
                            Set .CellPicture = imglstIcon.ListImages(2).Picture
                        End If
                    End If
                    
                    
                    AddExamDet .Row
                    
                    If Len(.TextMatrix(.Rows - 1, 0)) > 0 Then
                        .Rows = .Rows + 1
                    End If
                    
                    .Row = .Rows - 1
                    
                    .Col = 0
                    .CellBackColor = vbWhite
                    .CellForeColor = vbBlack
                    Set .CellPicture = imglstIcon.ListImages(6).Picture
                    .CellPictureAlignment = 4
                    .Col = 11
                    .CellFontName = "Wingdings"
                    .CellFontBold = True
                    .CellFontSize = 12
                    .CellAlignment = 4
                    .TextMatrix(.Row, .Col) = Chr(111)
                    .Col = 1
                    
                    .RowHeight(.Row) = 350
                    txtmsfgExam.Text = ""
                    .CellAlignment = 0
                    
                    If .Rows > 20 Then
                        VScrollExam.Visible = False
                    Else
                        VScrollExam.Visible = True
                    End If
                    .TopRow = .Rows - 1
                End With
            Else
                MsgBox "Unable to Add " & KP_arryExamSelected(1)
                txtmsfgExam.SetFocus
                highLight
                highLight_bg
            End If
        End If
        txtmsfgExam.SetFocus
        KeyAscii = 0
    ElseIf KeyAscii = 9 Then
        msfgExam.SetFocus
        msfgExam_EnterCell
        KeyAscii = 0
    Else
        If KeyAscii = Asc("'") Then
            KeyAscii = 0
        ElseIf KeyAscii = Asc("|") Then
            KeyAscii = 0
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
End Sub

Private Sub txtmsfgExam_LostFocus()
    txtmsfgExam.BackColor = vbWhite
End Sub

Private Sub txtmsfgResult_GotFocus()
    highLight
    highLight_bg
End Sub

Private Function setLenTxtBox(columnName As String) As Integer
    Select Case columnName
        Case "PID"
            setLenTxtBox = 10
        Case "PHOSTPID"
            setLenTxtBox = 10
        Case "PLNAME"
            setLenTxtBox = 50
        Case "PFNAME"
            setLenTxtBox = 50
        Case "PMNAME"
            setLenTxtBox = 50
        Case "PGENDER"
            setLenTxtBox = 10
        Case "PDOB"
            setLenTxtBox = 10
        Case "PCIVILSTATUS"
            setLenTxtBox = 10
        Case "PPHICID"
            setLenTxtBox = 20
        Case "PADDRESS1"
            setLenTxtBox = 50
        Case "PADDRESS2"
            setLenTxtBox = 50
        Case "PEMAIL"
            setLenTxtBox = 50
        Case "PPHONE"
            setLenTxtBox = 20
        Case "RequestDate"
            setLenTxtBox = 20
        Case "RequestWard"
            setLenTxtBox = 25
        Case "RequestRoom"
            setLenTxtBox = 25
        Case "RequestPType"
            setLenTxtBox = 10
        Case "RequestLocation"
            setLenTxtBox = 20
        Case "RequestID"
            setLenTxtBox = 10
        Case "RequestPID"
            setLenTxtBox = 20
        Case "RequestBy"
            setLenTxtBox = 10
        Case "RequestAge"
            setLenTxtBox = 20
        Case "RequestPriority"
            setLenTxtBox = 10
        Case "RequestComment"
            setLenTxtBox = 150
        Case "RequestStatus"
            setLenTxtBox = 10
        Case "RequestPhysicianID"
            setLenTxtBox = 60
        Case "Reserve1"
            setLenTxtBox = 20
        Case "Reserve2"
            setLenTxtBox = 20
        Case "Reserve3"
            setLenTxtBox = 20
        Case "Reserve4"
            setLenTxtBox = 20
        Case "Reserve5"
            setLenTxtBox = 20
    End Select
End Function

Private Sub setListcmbmsfgReqDet(strFieldName As String)
    
    Dim strSelectionCriteria
    Dim intRowSelection As Integer
    Dim rsSelection As New ADODB.Recordset

    strSQLStatement = "EXEC sp_tbSelectionValues_select '" & msfgResult.TextMatrix(msfgResult.Row, 7) & "'"
    Set rsSelection = myConnection.Execute(strSQLStatement)
    
    With cmbmsfgResult
        If .ListCount > 0 Then
            For intRowSelection = .ListCount To 1 Step -1
                .RemoveItem (0)
            Next
        End If
        If rsSelection.RecordCount > 0 Then
            For intRowSelection = 0 To rsSelection.RecordCount - 1
                .AddItem (rsSelection.Fields(1))
                rsSelection.MoveNext
            Next
        End If
    End With
    
End Sub

Private Sub txtmsfgResult_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim msfgResultCurrentRow As Integer
    Dim msfgResultNewRow As Integer
    Dim strSelectionValue As String

    If msfgResult.Col = 1 Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then

            msfgResultCurrentRow = msfgResult.Row
            msfgResult_LeaveCell

            With msfgResult
                If UCase(.TextMatrix(.Row, 6)) = "SELECTION" And .TextMatrix(.Row, 3) = "1" And .TextMatrix(.Row, 4) = "1" Then
                    If .TextMatrix(.Row, 2) = "Ward" Then
                        If Len(.TextMatrix(.Row, 10)) = 0 Then
                            If Len(Trim(txtmsfgResult.Text)) > 0 Then
                                strSelectionValue = showfrmSelection("2", txtmsfgResult.Text, False)
                                If Len(strSelectionValue) > 0 And strSelectionValue <> "CANCELED" Then
                                    .TextMatrix(.Row, 1) = getValue(strSelectionValue, 2)
                                    .TextMatrix(.Row, 10) = getValue(strSelectionValue, 0)
                                    .TextMatrix(21, 1) = ""
                                    .TextMatrix(21, 10) = ""
                                    .TextMatrix(22, 1) = ""
                                    .TextMatrix(23, 1) = getValue(strSelectionValue, 4)
                                    .TextMatrix(23, 10) = getValue(strSelectionValue, 3)
                                Else
                                    If Len(strSelectionValue) = 0 Then
                                        MsgBox "No record found", vbInformation + vbOKOnly, "LMS"
                                    End If
                                    .TextMatrix(.Row, 1) = ""
                                    .TextMatrix(.Row, 10) = ""
                                    If Len(.TextMatrix(21, 10)) = 0 Then
                                        .TextMatrix(23, 1) = ""
                                        .TextMatrix(23, 10) = ""
                                    End If
                                End If
                            Else
                                .TextMatrix(.Row, 1) = ""
                                .TextMatrix(.Row, 10) = ""
                                If Len(.TextMatrix(21, 10)) = 0 Then
                                    .TextMatrix(23, 1) = ""
                                    .TextMatrix(23, 10) = ""
                                End If
                            End If
                        End If
                        txtmsfgResult.Text = .TextMatrix(.Row, 1)
                    ElseIf .TextMatrix(.Row, 2) = "Room" Then
                        If Len(.TextMatrix(.Row, 10)) = 0 Then
                            If Len(Trim(txtmsfgResult.Text)) > 0 Then
                                strSelectionValue = showfrmSelection("4", txtmsfgResult.Text, False)
                                If Len(strSelectionValue) > 0 And strSelectionValue <> "CANCELED" Then
                                    .TextMatrix(.Row, 1) = getValue(strSelectionValue, 2)
                                    .TextMatrix(.Row, 10) = getValue(strSelectionValue, 0)
                                    .TextMatrix(20, 1) = ""
                                    .TextMatrix(20, 10) = ""
                                    .TextMatrix(22, 1) = ""
                                    .TextMatrix(23, 1) = getValue(strSelectionValue, 4)
                                    .TextMatrix(23, 10) = getValue(strSelectionValue, 3)
                                Else
                                    If Len(strSelectionValue) = 0 Then
                                        MsgBox "No record found", vbInformation + vbOKOnly, "LMS"
                                    End If
                                    .TextMatrix(.Row, 1) = ""
                                    .TextMatrix(.Row, 10) = ""
                                    If Len(.TextMatrix(20, 10)) = 0 Then
                                        .TextMatrix(23, 1) = ""
                                        .TextMatrix(23, 10) = ""
                                    End If
                                End If
                            Else
                                .TextMatrix(.Row, 1) = ""
                                .TextMatrix(.Row, 10) = ""
                                If Len(.TextMatrix(20, 10)) = 0 Then
                                    .TextMatrix(23, 1) = ""
                                    .TextMatrix(23, 10) = ""
                                End If
                            End If
                        End If
                        txtmsfgResult.Text = .TextMatrix(.Row, 1)
                    ElseIf .TextMatrix(.Row, 2) = "ReqRate" Then
                        If Len(.TextMatrix(.Row, 10)) = 0 Then
                            If Len(.TextMatrix(20, 10)) > 0 Then
                                If Len(txtmsfgResult.Text) > 0 Then
                                    strSelectionValue = showfrmSelection("5", .TextMatrix(20, 10), False, txtmsfgResult.Text)
                                    If Len(strSelectionValue) > 0 And strSelectionValue <> "CANCELED" Then
                                        .TextMatrix(.Row, 1) = getValue(strSelectionValue, 1)
                                        .TextMatrix(.Row, 10) = getValue(strSelectionValue, 0)
                                    Else
                                        If Len(strSelectionValue) > 0 Then
                                            MsgBox "No record found", vbInformation + vbOKOnly, "LMS"
                                        End If
                                        .TextMatrix(.Row, 1) = ""
                                        .TextMatrix(.Row, 10) = ""
                                    End If
                                End If
                            ElseIf Len(.TextMatrix(21, 10)) > 0 Then
                                If Len(txtmsfgResult.Text) > 0 Then
                                    strSelectionValue = showfrmSelection("6", .TextMatrix(21, 10), False, txtmsfgResult.Text)
                                    If Len(strSelectionValue) > 0 And strSelectionValue <> "CANCELED" Then
                                        .TextMatrix(.Row, 1) = getValue(strSelectionValue, 1)
                                        .TextMatrix(.Row, 10) = getValue(strSelectionValue, 0)
                                    Else
                                        If Len(strSelectionValue) = 0 Then
                                            MsgBox "No record found", vbInformation + vbOKOnly, "LMS"
                                        End If
                                        .TextMatrix(.Row, 1) = ""
                                        .TextMatrix(.Row, 10) = ""
                                    End If
                                End If
                            Else
                                .TextMatrix(.Row, 1) = ""
                                .TextMatrix(.Row, 10) = ""
                            End If
                        End If
                        txtmsfgResult.Text = .TextMatrix(.Row, 1)
                    Else
                        If Len(Trim(txtmsfgResult.Text)) > 0 Then
                            strSelectionValue = SetFieldWithSelection(.TextMatrix(.Row, 7), txtmsfgResult.Text)
                            If Len(strSelectionValue) > 0 Then
                                .TextMatrix(.Row, 1) = getValue(strSelectionValue, 1)
                                .TextMatrix(.Row, 10) = getValue(strSelectionValue, 0)
                            Else
                                .TextMatrix(.Row, 1) = ""
                                .TextMatrix(.Row, 10) = ""
                            End If
                        End If
                    
'                        If Len(txtmsfgResult.Text) > 0 Then
'                            strSelectionValue = showfrmSelection(.TextMatrix(.Row, 7), txtmsfgResult.Text, False)
'                            If Len(strSelectionValue) > 0 And strSelectionValue <> "CANCELED" Then
'                                .TextMatrix(.Row, 1) = getValue(strSelectionValue, 1)
'                                .TextMatrix(.Row, 10) = getValue(strSelectionValue, 0)
'                            Else
'                                If Len(strSelectionValue) = 0 Then
'                                    MsgBox "No record found", vbInformation + vbOKOnly, "LMS"
'                                End If
'                                .TextMatrix(.Row, 1) = ""
'                                .TextMatrix(.Row, 10) = ""
'                            End If
'                        End If
                        
                        txtmsfgResult.Text = .TextMatrix(.Row, 1)
                    End If
                End If
                If KeyCode = vbKeyUp Then
                    For msfgResultNewRow = .Row - 1 To 0 Step -1
                        If .TextMatrix(msfgResultNewRow, 3) = "1" And .TextMatrix(msfgResultNewRow, 4) = "1" Then
                            .Row = msfgResultNewRow
                            Exit For
                        End If
                    Next
                Else
                    For msfgResultNewRow = .Row + 1 To .Rows - 1
                        If .TextMatrix(msfgResultNewRow, 3) = "1" And .TextMatrix(msfgResultNewRow, 4) = "1" Then
                            .Row = msfgResultNewRow
                            Exit For
                        End If
                    Next
                End If
                
                If msfgResult.Row = msfgResultCurrentRow Then
                    msfgExam.Col = 1
                    msfgExam_EnterCell
                Else
                    msfgResult_EnterCell
                    highLight
                End If
                
                If KeyCode = vbKeyUp Then
                    If .RowIsVisible(3) Then
                        .TopRow = 0
                    End If
                End If
            End With
        ElseIf KeyCode = 113 And msfgResult.TextMatrix(msfgResult.Row, 6) = "SELECTION" Then
            If msfgResult.TextMatrix(msfgResult.Row, 2) = "Ward" Then
                With msfgResult
                    strSelectionValue = showfrmSelection("2", txtmsfgResult.Text, True)
                    If Not isExiting Then
                        If strSelectionValue <> "CANCELED" Then
                            .TextMatrix(.Row, 1) = getValue(strSelectionValue, 2)
                            .TextMatrix(.Row, 10) = getValue(strSelectionValue, 0)
                            .TextMatrix(21, 1) = ""
                            .TextMatrix(21, 10) = ""
                            .TextMatrix(22, 1) = ""
                            .TextMatrix(23, 1) = getValue(strSelectionValue, 4)
                            .TextMatrix(23, 10) = getValue(strSelectionValue, 3)
                        End If
                        txtmsfgResult.Text = .TextMatrix(.Row, 1)
                    End If
                End With
            ElseIf msfgResult.TextMatrix(msfgResult.Row, 2) = "Room" Then
                With msfgResult
                    strSelectionValue = showfrmSelection("4", txtmsfgResult.Text, True)
                    If Not isExiting Then
                        If strSelectionValue <> "CANCELED" Then
                            .TextMatrix(.Row, 1) = getValue(strSelectionValue, 2)
                            .TextMatrix(.Row, 10) = getValue(strSelectionValue, 0)
                            .TextMatrix(20, 1) = ""
                            .TextMatrix(20, 10) = ""
                            .TextMatrix(22, 1) = ""
                            .TextMatrix(23, 1) = getValue(strSelectionValue, 4)
                            .TextMatrix(23, 10) = getValue(strSelectionValue, 3)
                        End If
                        txtmsfgResult.Text = .TextMatrix(.Row, 1)
                    End If
                End With
            ElseIf msfgResult.TextMatrix(msfgResult.Row, 2) = "ReqRate" Then
                With msfgResult
                    If Len(.TextMatrix(20, 10)) > 0 Then
                        strSelectionValue = showfrmSelection("5", .TextMatrix(20, 10), True, txtmsfgResult.Text)
                        If Not isExiting Then
                            If strSelectionValue <> "CANCELED" Then
                                .TextMatrix(.Row, 1) = getValue(strSelectionValue, 1)
                                .TextMatrix(.Row, 10) = getValue(strSelectionValue, 0)
                            End If
                        End If
                    ElseIf Len(.TextMatrix(21, 10)) > 0 Then
                        strSelectionValue = showfrmSelection("6", .TextMatrix(21, 10), True, txtmsfgResult.Text)
                        If Not isExiting Then
                            If strSelectionValue <> "CANCELED" Then
                               .TextMatrix(.Row, 1) = getValue(strSelectionValue, 1)
                               .TextMatrix(.Row, 10) = getValue(strSelectionValue, 0)
                            End If
                        End If
                    End If
                    txtmsfgResult.Text = .TextMatrix(.Row, 1)
                End With
            Else
                strSelectionValue = searchObj.showSearchCode(msfgResult.TextMatrix(msfgResult.Row, 7), txtmsfgResult.Text, False, , True, True)
                If Not isExiting Then
                    msfgResult_EnterCell
                    With msfgResult
                        If Len(Trim(strSelectionValue)) > 0 Then
                            .TextMatrix(.Row, 10) = getValue(strSelectionValue, 0)
                            .TextMatrix(.Row, 1) = getValue(strSelectionValue, 1)
                            txtmsfgResult.Text = Mid(strSelectionValue, InStr(1, strSelectionValue, "|") + 1)
                        Else
                            If Len(.TextMatrix(.Row, 1)) = 0 And Len(.TextMatrix(.Row, 10)) = 0 Then
                                .TextMatrix(.Row, 1) = ""
                                .TextMatrix(.Row, 10) = ""
                            End If
                        End If
                    End With
                End If
                    
            
'                With msfgResult
'                    strSelectionValue = showfrmSelection(.TextMatrix(.Row, 7), txtmsfgResult.Text, True)
'                    If Not isExiting Then
'                        If strSelectionValue <> "CANCELED" Then
'                            .TextMatrix(.Row, 1) = getValue(strSelectionValue, 1)
'                            .TextMatrix(.Row, 10) = getValue(strSelectionValue, 0)
'                        End If
'                    End If
'                    txtmsfgResult.Text = .TextMatrix(.Row, 1)
'                End With
            End If
            With msfgResult
                For msfgResultNewRow = .Row + 1 To .Rows - 1
                    If .TextMatrix(msfgResultNewRow, 3) = "1" And .TextMatrix(msfgResultNewRow, 4) = "1" Then
                        .Row = msfgResultNewRow
                        Exit For
                    End If
                Next
            End With
            If msfgResult.Row = msfgResultCurrentRow Then
                msfgExam.Col = 1
                msfgExam_EnterCell
            Else
                msfgResult_EnterCell
                highLight
            End If
        End If
    End If
    
End Sub

Private Sub txtmsfgResult_KeyPress(KeyAscii As Integer)

    Dim msfgResultCurrentRow As Integer
    Dim msfgResultNewRow As Integer
    Dim strSelectionValue As String

    If KeyAscii = 13 Or KeyAscii = 9 Then
        msfgResultCurrentRow = msfgResult.Row
        msfgResult_LeaveCell
        With msfgResult
            If UCase(.TextMatrix(.Row, 6)) = "SELECTION" And .TextMatrix(.Row, 3) = "1" And .TextMatrix(.Row, 4) = "1" Then
                If .TextMatrix(.Row, 2) = "Ward" Then
                    If Len(.TextMatrix(.Row, 10)) = 0 Then
                        If Len(Trim(txtmsfgResult.Text)) > 0 Then
                            strSelectionValue = showfrmSelection("2", txtmsfgResult.Text, False)
                            If Len(strSelectionValue) > 0 And strSelectionValue <> "CANCELED" Then
                                .TextMatrix(.Row, 1) = getValue(strSelectionValue, 2)
                                .TextMatrix(.Row, 10) = getValue(strSelectionValue, 0)
                                .TextMatrix(21, 1) = ""
                                .TextMatrix(21, 10) = ""
                                .TextMatrix(22, 1) = ""
                                .TextMatrix(23, 1) = getValue(strSelectionValue, 4)
                                .TextMatrix(23, 10) = getValue(strSelectionValue, 3)
                            Else
                                If Len(strSelectionValue) = 0 Then
                                    MsgBox "No record found", vbInformation + vbOKOnly, "LMS"
                                End If
                                .TextMatrix(.Row, 1) = ""
                                .TextMatrix(.Row, 10) = ""
                                If Len(.TextMatrix(21, 10)) = 0 Then
                                    .TextMatrix(23, 1) = ""
                                    .TextMatrix(23, 10) = ""
                                End If
                            End If
                        Else
                            .TextMatrix(.Row, 1) = ""
                            .TextMatrix(.Row, 10) = ""
                            If Len(.TextMatrix(21, 10)) = 0 Then
                                .TextMatrix(23, 1) = ""
                                .TextMatrix(23, 10) = ""
                            End If
                        End If
                    End If
                    txtmsfgResult.Text = .TextMatrix(.Row, 1)
                ElseIf .TextMatrix(.Row, 2) = "Room" Then
                    If Len(.TextMatrix(.Row, 10)) = 0 Then
                        If Len(Trim(txtmsfgResult.Text)) > 0 Then
                            strSelectionValue = showfrmSelection("4", txtmsfgResult.Text, False)
                            If Len(strSelectionValue) > 0 And strSelectionValue <> "CANCELED" Then
                                .TextMatrix(.Row, 1) = getValue(strSelectionValue, 2)
                                .TextMatrix(.Row, 10) = getValue(strSelectionValue, 0)
                                .TextMatrix(20, 1) = ""
                                .TextMatrix(20, 10) = ""
                                .TextMatrix(22, 1) = ""
                                .TextMatrix(23, 1) = getValue(strSelectionValue, 4)
                                .TextMatrix(23, 10) = getValue(strSelectionValue, 3)
                            Else
                                If Len(strSelectionValue) = 0 Then
                                    MsgBox "No record found", vbInformation + vbOKOnly, "LMS"
                                End If
                                .TextMatrix(.Row, 1) = ""
                                .TextMatrix(.Row, 10) = ""
                                If Len(.TextMatrix(20, 10)) = 0 Then
                                    .TextMatrix(23, 1) = ""
                                    .TextMatrix(23, 10) = ""
                                End If
                            End If
                        Else
                            .TextMatrix(.Row, 1) = ""
                            .TextMatrix(.Row, 10) = ""
                            If Len(.TextMatrix(20, 10)) = 0 Then
                                .TextMatrix(23, 1) = ""
                                .TextMatrix(23, 10) = ""
                            End If
                        End If
                    End If
                    txtmsfgResult.Text = .TextMatrix(.Row, 1)
                ElseIf .TextMatrix(.Row, 2) = "ReqRate" Then
                    If Len(.TextMatrix(.Row, 10)) = 0 Then
                        If Len(.TextMatrix(20, 10)) > 0 Then
                            If Len(txtmsfgResult.Text) > 0 Then
                                strSelectionValue = showfrmSelection("5", .TextMatrix(20, 10), False, txtmsfgResult.Text)
                                If Len(strSelectionValue) > 0 And strSelectionValue <> "CANCELED" Then
                                    .TextMatrix(.Row, 1) = getValue(strSelectionValue, 1)
                                    .TextMatrix(.Row, 10) = getValue(strSelectionValue, 0)
                                Else
                                    If Len(strSelectionValue) = 0 Then
                                        MsgBox "No record found", vbInformation + vbOKOnly, "LMS"
                                    End If
                                    .TextMatrix(.Row, 1) = ""
                                    .TextMatrix(.Row, 10) = ""
                                End If
                            End If
                        ElseIf Len(.TextMatrix(21, 10)) > 0 Then
                            If Len(txtmsfgResult.Text) > 0 Then
                                strSelectionValue = showfrmSelection("6", .TextMatrix(21, 10), False, txtmsfgResult.Text)
                                If Len(strSelectionValue) > 0 And strSelectionValue <> "CANCELED" Then
                                    .TextMatrix(.Row, 1) = getValue(strSelectionValue, 1)
                                    .TextMatrix(.Row, 10) = getValue(strSelectionValue, 0)
                                Else
                                    If Len(strSelectionValue) = 0 Then
                                        MsgBox "No record found", vbInformation + vbOKOnly, "LMS"
                                    End If
                                    .TextMatrix(.Row, 1) = ""
                                    .TextMatrix(.Row, 10) = ""
                                End If
                            End If
                        Else
                            .TextMatrix(.Row, 1) = ""
                            .TextMatrix(.Row, 10) = ""
                        End If
                    End If
                    txtmsfgResult.Text = .TextMatrix(.Row, 1)
                Else
                    If Len(txtmsfgResult.Text) > 0 Then
                        strSelectionValue = SetFieldWithSelection(.TextMatrix(.Row, 7), txtmsfgResult.Text)
                        If Len(strSelectionValue) > 0 Then
                            .TextMatrix(.Row, 1) = getValue(strSelectionValue, 1)
                            txtmsfgResult.Text = .TextMatrix(.Row, 1)
                            .TextMatrix(.Row, 10) = getValue(strSelectionValue, 0)
                        Else
                            .TextMatrix(.Row, 1) = ""
                            .TextMatrix(.Row, 10) = ""
                        End If
                    
'                        strSelectionValue = showfrmSelection(.TextMatrix(.Row, 7), txtmsfgResult.Text, False)
'                        If Len(strSelectionValue) > 0 And strSelectionValue <> "CANCELED" Then
'                            .TextMatrix(.Row, 1) = getValue(strSelectionValue, 1)
'                            .TextMatrix(.Row, 10) = getValue(strSelectionValue, 0)
'                        Else
'                            If Len(strSelectionValue) = 0 Then
'                                MsgBox "No record found", vbInformation + vbOKOnly, "LMS"
'                            End If
'                            .TextMatrix(.Row, 1) = ""
'                            .TextMatrix(.Row, 10) = ""
'                        End If
                    End If
                    txtmsfgResult.Text = .TextMatrix(.Row, 1)
                End If
            End If
            For msfgResultNewRow = .Row + 1 To .Rows - 1
                If .TextMatrix(msfgResultNewRow, 3) = "1" And .TextMatrix(msfgResultNewRow, 4) = "1" Then
                    .Row = msfgResultNewRow
                    Exit For
                End If
            Next
        End With
    
        If msfgResult.Row = msfgResultCurrentRow Then
            txtmsfgExam.SetFocus
        Else
            msfgResult_EnterCell
            highLight
        End If
        
        KeyAscii = 0
    ElseIf KeyAscii = Asc("|") Or KeyAscii = Asc("'") Or KeyAscii = Asc("""") Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        msfgResult.TextMatrix(msfgResult.Row, 1) = ""
        msfgResult.TextMatrix(msfgResult.Row, 10) = ""
    End If
    
    If msfgResult.TextMatrix(msfgResult.Row, 2) = "Gender" Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If InStr(1, "MF", Chr(KeyAscii)) > 0 Then
            If Chr(KeyAscii) = "M" Then txtmsfgResult.Text = "MALE"
            If Chr(KeyAscii) = "F" Then txtmsfgResult.Text = "FEMALE"
        End If
        If KeyAscii = 32 Then
            If txtmsfgResult.Text = "MALE" Then
                txtmsfgResult.Text = "FEMALE"
            ElseIf txtmsfgResult.Text = "FEMALE" Then
                txtmsfgResult.Text = "MALE"
            Else
                txtmsfgResult.Text = "MALE"
            End If
        ElseIf KeyAscii = 8 Then
            txtmsfgResult.Text = ""
        End If
        KeyAscii = 0
    End If
    
End Sub

Private Sub setmsfgExam()
    
    Dim rsmsfgExam As New ADODB.Recordset
    Dim intmsfgExamColCtr As Integer
    Dim intmsfgExamRowCtr As Integer
    
    With msfgExam
        .Visible = False
        
        If Len(Trim(edtRequestID)) = 0 Then
            strSQLStatement = "EXEC sp_tbLabRequestValues_select 'INIT'"
            Set rsmsfgExam = myConnection.Execute(strSQLStatement)
            Set .DataSource = rsmsfgExam
            .Refresh
            blnNoExistingTest = True
        Else
            strSQLStatement = "EXEC sp_tbLabRequestValues_select '" & edtRequestID & "'"
            Set rsmsfgExam = myConnection.Execute(strSQLStatement)
            rsmsfgExam.Filter = "RVSectionID ='" & edtSectionID & "' AND Specimen = '" & edtSpecimen & "'" & _
                                    " AND RVSampleID ='" & edtSampleNumber & "'"
            If rsmsfgExam.RecordCount > 0 Then
                Set .DataSource = rsmsfgExam
                .Refresh
                .Rows = .Rows + 1
                blnNoExistingTest = False
            Else
                blnNoExistingTest = True
            End If
        End If
        
        .RowHeight(0) = 400
        For intmsfgExamColCtr = 1 To .Cols - 1
            .TextArray(intmsfgExamColCtr) = rsmsfgExam.Fields(intmsfgExamColCtr - 1).Name
            Debug.Print (rsmsfgExam.Fields(intmsfgExamColCtr - 1).Name)
            .ColWidth(intmsfgExamColCtr) = 0
        Next
            
        For intmsfgExamRowCtr = 1 To .Rows - 1
            .Row = intmsfgExamRowCtr
            .Col = 0
            .CellAlignment = 0
            .CellBackColor = vbWhite
            'ICON''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Select Case .TextMatrix(.Row, 6)
                Case ""
                    Set .CellPicture = imglstIcon.ListImages(6).Picture
                Case "0"
                    Set .CellPicture = imglstIcon.ListImages(6).Picture
                Case "1"
                    Set .CellPicture = imglstIcon.ListImages(2).Picture
                Case "2"
                    Set .CellPicture = imglstIcon.ListImages(3).Picture
                    blnFormIsLocked = lockForm("Locked. Sample Process Ongoing")
                Case "3"
                    Set .CellPicture = imglstIcon.ListImages(4).Picture
                Case "4"
                    Set .CellPicture = imglstIcon.ListImages(8).Picture
                Case "5"
                    Set .CellPicture = imglstIcon.ListImages(7).Picture
                    blnFormIsLocked = lockForm("Locked. Request Approved")
                Case "6"
                    Set .CellPicture = imglstIcon.ListImages(9).Picture
            End Select
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            .CellPictureAlignment = 4
            
            'STAT INDICATOR'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            .Col = 11
'            .CellFontName = "Wingdings"
            .CellFontBold = True
            .CellFontSize = 12
            .CellAlignment = 4
            Select Case .TextMatrix(.Row, 11)
                Case "1"
                    .TextMatrix(.Row, .Col) = Chr(254)
                    .TextMatrix(.Row, 13) = Chr(254)
                Case Else
                    .TextMatrix(.Row, .Col) = Chr(111)
                    .TextMatrix(.Row, 13) = Chr(111)
            End Select
            
            .RowHeight(.Row) = 350
        Next
    
        .TextArray(0) = "Status"
        .TextArray(1) = "Exam Code"
        .TextArray(2) = "Exam Name"
        .TextArray(3) = "Specimen"
        .TextArray(4) = "Section"
        .TextArray(5) = "Amount"
        
        .TextArray(11) = "      STAT"
       
        .ColWidth(0) = 800
        .ColWidth(1) = 1200
        .ColWidth(2) = 4120
        .ColWidth(3) = 2000
        .ColWidth(4) = 2000
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 1150
        .ColWidth(12) = 0
        If blnNoExistingTest = False Then
            setmsfgExamDet
        End If
        
        .Visible = True
    End With

End Sub

Private Sub AddExamDet(introw As Integer)
    
    Dim rsExamDetails As New ADODB.Recordset
    Dim RowCtr As Integer
    
    With msfgExam
    
        strSQLStatement = "EXEC sp_getParamProfileExam_Select '" & .TextMatrix(introw, 10) & "'," & _
                                    "'" & .TextMatrix(introw, 7) & "'," & _
                                    "'" & .TextMatrix(introw, 1) & "'," & _
                                    "'" & .TextMatrix(introw, 3) & "'"
                                    
        Set rsExamDetails = myConnection.Execute(strSQLStatement)
        
        If rsExamDetails.RecordCount > 0 Then
            For RowCtr = 0 To rsExamDetails.RecordCount - 1
                msfgExamDet.TextMatrix(msfgExamDet.Rows - 1, 0) = rsExamDetails.Fields(0)
                msfgExamDet.TextMatrix(msfgExamDet.Rows - 1, 1) = rsExamDetails.Fields(1)
                msfgExamDet.TextMatrix(msfgExamDet.Rows - 1, 2) = rsExamDetails.Fields(2)
                msfgExamDet.TextMatrix(msfgExamDet.Rows - 1, 3) = rsExamDetails.Fields(3)
                msfgExamDet.TextMatrix(msfgExamDet.Rows - 1, 4) = rsExamDetails.Fields(4)
                msfgExamDet.TextMatrix(msfgExamDet.Rows - 1, 5) = rsExamDetails.Fields(5)
                rsExamDetails.MoveNext
                msfgExamDet.Rows = msfgExamDet.Rows + 1
            Next
        End If
            
    End With
    
End Sub

Private Sub setmsfgExamDet()
    Dim rsAllExamDetails As New ADODB.Recordset
    Dim RowAllExamCtr As Integer
    
    With msfgExamDet
        strSQLStatement = "EXEC sp_GetOtherExam_select '" & edtRequestID & "'"
        'Debug.Print rsAllExamDetails.RecordCount
        Set rsAllExamDetails = myConnection.Execute(strSQLStatement)
        
    
        If rsAllExamDetails.RecordCount > 0 Then
            For RowAllExamCtr = 0 To rsAllExamDetails.RecordCount - 1
                .TextMatrix(.Rows - 1, 0) = rsAllExamDetails.Fields("RVSectionID").Value
                .TextMatrix(.Rows - 1, 1) = rsAllExamDetails.Fields("RVTestID").Value
                .TextMatrix(.Rows - 1, 2) = rsAllExamDetails.Fields("Specimen").Value
                .TextMatrix(.Rows - 1, 3) = rsAllExamDetails.Fields("LabExamCode").Value
                .TextMatrix(.Rows - 1, 4) = rsAllExamDetails.Fields("InstrumentRequestCode").Value
                .TextMatrix(.Rows - 1, 5) = rsAllExamDetails.Fields("LabExamType").Value
                rsAllExamDetails.MoveNext
                .Rows = .Rows + 1
            Next
        End If
    End With
    
End Sub

Private Function blnExamOccured(examCriteria As String) As Boolean
    
    Dim tempSectionID As String
    Dim tempExamCode As String
    Dim tempExamType As String
    Dim tempSpecimen As String
    Dim rsCheckParam As New ADODB.Recordset
    Dim msfgExamDetRowCtr As Integer
    Dim rsCheckParamRowCtr As Integer
    Dim EO_arryExamSelected() As String
    
    EO_arryExamSelected = Split(examCriteria, "|")
    
    tempExamType = EO_arryExamSelected(3)
    tempSectionID = EO_arryExamSelected(4)
    tempExamCode = EO_arryExamSelected(0)
    tempSpecimen = EO_arryExamSelected(2)

    strSQLStatement = "EXEC sp_getParamProfileExam_Select '" & tempExamType & "','" & tempSectionID & "'," & _
                        "'" & tempExamCode & "','" & tempSpecimen & "'"
    
    Set rsCheckParam = myConnection.Execute(strSQLStatement)
    
    With msfgExamDet
        For rsCheckParamRowCtr = 0 To rsCheckParam.RecordCount - 1
            
            For msfgExamDetRowCtr = 1 To .Rows - 1
                If RTrim(rsCheckParam.Fields(5).Value) = "P" Then
                    If RTrim(rsCheckParam.Fields(0).Value) = RTrim(.TextMatrix(msfgExamDetRowCtr, 0)) And _
                        RTrim(rsCheckParam.Fields(1).Value) = RTrim(.TextMatrix(msfgExamDetRowCtr, 1)) And _
                        RTrim(rsCheckParam.Fields(2).Value) = RTrim(.TextMatrix(msfgExamDetRowCtr, 2)) And _
                        RTrim(rsCheckParam.Fields(3).Value) = RTrim(.TextMatrix(msfgExamDetRowCtr, 3)) And _
                        RTrim(rsCheckParam.Fields(4).Value) = RTrim(.TextMatrix(msfgExamDetRowCtr, 4)) And _
                        RTrim(rsCheckParam.Fields(5).Value) = RTrim(.TextMatrix(msfgExamDetRowCtr, 5)) Then
                        blnExamOccured = True
                        Exit For
                    End If
                Else
                    If RTrim(rsCheckParam.Fields(0).Value) = RTrim(.TextMatrix(msfgExamDetRowCtr, 0)) And _
                        RTrim(rsCheckParam.Fields(2).Value) = RTrim(.TextMatrix(msfgExamDetRowCtr, 2)) And _
                        RTrim(rsCheckParam.Fields(3).Value) = RTrim(.TextMatrix(msfgExamDetRowCtr, 3)) And _
                        RTrim(rsCheckParam.Fields(4).Value) = RTrim(.TextMatrix(msfgExamDetRowCtr, 4)) And _
                        RTrim(rsCheckParam.Fields(5).Value) = RTrim(.TextMatrix(msfgExamDetRowCtr, 5)) Then
                        blnExamOccured = True
                        Exit For
                    End If
                End If
            Next
            If blnExamOccured Then
                Exit For
            Else
                rsCheckParam.MoveNext
            End If
        Next
    End With
    
End Function
Private Sub DelExamDet(reqSectionID As Integer, reqExamCode As Integer)
    Dim delRowCtr As Integer
    With msfgExamDet
        delRowCtr = 1
        Do
            If .TextMatrix(delRowCtr, 0) = reqSectionID And .TextMatrix(delRowCtr, 1) = reqExamCode Then
                .RemoveItem (delRowCtr)
            Else
                delRowCtr = delRowCtr + 1
            End If
        Loop Until delRowCtr = .Rows - 1
    End With

End Sub

Private Function ReqIsLocked() As Boolean
    
    Dim toReturn As Boolean
    Dim rsCheckReqLocked As New ADODB.Recordset
    
    strSQLStatement = "EXEC sp_LockRequest_Select '" & intWorkstation & "','" & edtRequestID & "'," & _
                        "'" & edtSpecimen & "'"
                        
    Set rsCheckReqLocked = myConnection.Execute(strSQLStatement)
    
    If rsCheckReqLocked.RecordCount > 0 Then
        toReturn = True
    Else
        toReturn = False
    End If
        
    ReqIsLocked = toReturn
    
End Function

Private Sub ReqLock()
        
    Dim rsLockRequest As New ADODB.Recordset
    
    If Len(Trim(edtRequestID)) > 0 Then
        strSQLStatement = "EXEC sp_LockRequest_Insert '" & intWorkstation & "','" & edtRequestID & "'," & _
                            "'" & edtSpecimen & "'"
                            
        Set rsLockRequest = myConnection.Execute(strSQLStatement)
    End If
    
End Sub

Private Sub ReqUnlock()
    
    Dim rsUnlockRequest As New ADODB.Recordset
    
    If Len(Trim(edtRequestID)) > 0 Then
        strSQLStatement = "EXEC sp_LockRequest_Delete '" & intWorkstation & "'"
        
        Set rsUnlockRequest = myConnection.Execute(strSQLStatement)
    End If
    
End Sub

Private Sub setIntialCursor()
    
    Dim intCtrInitCurs As Integer
    Dim intRowNum As Integer
    Dim intInitCmbSelection As Integer
    
    With msfgResult
        
        If ENA_START_CUR = "1" Then
            For intCtrInitCurs = 0 To .Rows - 1
                If .TextMatrix(intCtrInitCurs, 9) = "1" Then
                    .Row = intCtrInitCurs
                    Exit For
                End If
            Next
        Else
            For intCtrInitCurs = 0 To .Rows - 1
                If Len(Trim(.TextMatrix(intCtrInitCurs, 1))) = 0 And .TextMatrix(intCtrInitCurs, 3) = "1" And _
                    .TextMatrix(intCtrInitCurs, 4) = "1" Then
                    .Row = intCtrInitCurs
                    Exit For
                End If
            Next
        End If
        
        If .TextMatrix(.Row, 3) = "1" And .TextMatrix(.Row, 4) = "1" Then
            .CellAlignment = 0
            If .TextMatrix(.Row, 6) = "TEXT" Or .TextMatrix(.Row, 6) = "SELECTION" Then
                txtmsfgResult.Move (.CellLeft + .Left), (.CellTop + .Top), .CellWidth, .CellHeight
                txtmsfgResult.Visible = True
                txtmsfgResult.Enabled = True
                txtmsfgResult.Text = .TextMatrix(.Row, .Col)
                txtmsfgResult.MaxLength = setLenTxtBox(.TextMatrix(.Row, 2))
                txtmsfgResult.TabStop = True
                txtmsfgResult.TabIndex = 0
            Else
                cmbmsfgResult.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
                cmbmsfgResult.Visible = True
                cmbmsfgResult.Enabled = True
                setListcmbmsfgReqDet .TextMatrix(.Row, 0)
                If Len(.TextMatrix(.Row, .Col)) > 0 Then
                    For intInitCmbSelection = 0 To cmbmsfgResult.ListCount - 1
                        If .TextMatrix(.Row, 1) = cmbmsfgResult.List(intInitCmbSelection) Then
                            cmbmsfgResult.ListIndex = intInitCmbSelection
                        End If
                    Next
                End If
                cmbmsfgResult.TabStop = False
                cmbmsfgResult.TabIndex = 0
            End If
        End If
        
    End With
    
End Sub

Private Function blnMandatoryClear() As Boolean
    Dim checkMandatory As Integer
    Dim strFieldMandatory As String
    
    With msfgResult
        For checkMandatory = 0 To .Rows - 1
            .TextMatrix(checkMandatory, 1) = Replace(.TextMatrix(checkMandatory, 1), "'", "''")
            If .TextMatrix(checkMandatory, 3) = "1" And .TextMatrix(checkMandatory, 4) = "1" Then
                If Len(Trim(.TextMatrix(checkMandatory, 1))) = 0 And .TextMatrix(checkMandatory, 5) = "1" Then
                    strFieldMandatory = strFieldMandatory & _
                                        IIf(strFieldMandatory = "", .TextMatrix(checkMandatory, 0), _
                                            ", " & .TextMatrix(checkMandatory, 0))
                End If
            End If
        Next
    End With
    
    If Len(strFieldMandatory) > 0 Then
        MsgBox strFieldMandatory & " are mandatory field/s", vbInformation + vbOKOnly, "LMS"
        blnMandatoryClear = False
    Else
        blnMandatoryClear = True
    End If
    
End Function

Private Sub sizeFix()
    Me.Width = Me.Width - 1000
    fraBody(0).Width = fraBody(0).Width - 1000
    With msfgExam
        .Width = .Width - 1000
        
        .ColWidth(0) = 600
        .ColWidth(1) = 1000
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColWidth(4) = 1500
        .ColWidth(5) = 4170
        
        lblLocked.Width = .Width
    End With

End Sub

Private Function lockForm(lockMessage As String) As Boolean
    txtmsfgResult.Visible = False
    txtmsfgResult.Enabled = False
    cmbmsfgResult.Visible = False
    cmbmsfgResult.Enabled = False
    txtmsfgExam.Enabled = False
    cmdExamMenu.Enabled = False
    txtmsfgExam.BackColor = &H8000000A
    msfgResult.Enabled = False
    msfgExam.Enabled = False
    msfgResult.BackColor = &HC0C0C0
    tlbarReq.Buttons(1).Enabled = False
    tlbarReq.Buttons(2).Enabled = False
    tlbarReq.Buttons(3).Enabled = False
    mnucmdSave.Enabled = False
    mnucmdPrntLbl.Enabled = False
    mnucmdRmvTest.Enabled = False
    lblText(1).Visible = False
    With txtRemarks
        .Enabled = False
        .Visible = False
        .BackColor = &H8000000A
    End With
    With lblLocked
        .Caption = vbCrLf & vbCr & "Locked. Sample is already opened by another user"
        .Top = 7890
        .Left = 90
        .Width = 16200
        .Height = 1725
        .Visible = True
    End With
    lockForm = True
End Function

Private Function SetFieldWithSelection(srchTransaction As String, srchItem As String) As String

    Dim rsSelection As New ADODB.Recordset
    Dim strTempReturn As String
    Dim ctrColConcat As Integer
    
    If IsNumeric(srchItem) Then
        strSQLStatement = "EXEC sp_SearchCodes '" & srchTransaction & "','" & srchItem & "',''"
    Else
        strSQLStatement = "EXEC sp_SearchCodes '" & srchTransaction & "','','" & Replace(srchItem, "'", "''") & "'"
    End If

    If msfgResult.Row = 27 Then
        Set rsSelection = myConnection.Execute(strSQLStatement)
        If rsSelection.RecordCount > 0 Then
            If rsSelection.RecordCount > 1 Then
                strTempReturn = searchObj.showSearchCode(srchTransaction, srchItem, False, , True, True)
            Else
                 For ctrColConcat = 0 To rsSelection.Fields.Count - 1
                    If Len(strTempReturn) = 0 Then
                        strTempReturn = rsSelection.Fields(ctrColConcat)
                    Else
                        strTempReturn = strTempReturn & "|" & rsSelection.Fields(ctrColConcat)
                    End If
                 Next
            End If
            SetFieldWithSelection = strTempReturn
        Else
            MsgBox "No Record Found...", vbOKOnly + vbInformation, "LMS"
            SetFieldWithSelection = ""
        End If
    Else
        With msfgResult
            SetFieldWithSelection = .TextMatrix(.Row, 10) & "|" & .TextMatrix(.Row, 1)
        End With
    End If
End Function

Private Sub delTempPatientDemo()
    Dim rsDelPDemo As New ADODB.Recordset
    
    strSQLStatement = "EXEC sp_PatientDemoTemp_delete " & intWorkstation
    Set rsDelPDemo = myConnection.Execute(strSQLStatement)
    
    frmSearchRequest.txtPatientID.Text = ""
End Sub

Private Sub txtRemarks_GotFocus()
    msfgResult_LeaveCell
    highLight_bg
End Sub

Private Sub txtRemarks_LostFocus()
    txtRemarks.BackColor = vbWhite
End Sub

Public Function RemoveFromTestToDel(TextID As String) As String
    Dim intmsfgToDelctr As Integer
    Dim RetDet As String
    '0 RequestID
    '1 SectionID
    '2 SampleID
    '3 TestID
    '4 Exam Name
    '5 Status
    '6 RVID
    '7 Priority
    '8 PanelID
            
    With msfgExamDel
        For intmsfgToDelctr = 1 To .Rows - 1
            .Row = intmsfgToDelctr
            If .TextMatrix(.Row, 3) = TextID Then
                RetDet = .TextMatrix(.Row, 5) & "|" & .TextMatrix(.Row, 6) & "|" & .TextMatrix(.Row, 7) & "|" & .TextMatrix(.Row, 8) & "|" & .TextMatrix(.Row, 1)
                .RemoveItem (.Row)
                Exit For
            End If
        Next
    End With
    
    RemoveFromTestToDel = RetDet
End Function

Private Function CheckWithSTAT() As String
    Dim ctrmsfgExamRow As Integer
    Dim retWithStat As String
    
    retWithStat = "0"
    With msfgExam
        For ctrmsfgExamRow = 1 To .Rows - 1
            If .TextMatrix(ctrmsfgExamRow, 11) = Chr(254) Then
                retWithStat = "1"
                Exit For
            End If
        Next
    End With

    CheckWithSTAT = retWithStat
    
End Function


