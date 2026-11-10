VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmRequest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Request Form"
   ClientHeight    =   10245
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   16620
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10245
   ScaleWidth      =   16620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   10335
      Index           =   0
      Left            =   -15
      TabIndex        =   8
      Top             =   -90
      Width           =   16635
      Begin VB.CommandButton cmdNewRequest 
         Caption         =   "&New Request"
         Enabled         =   0   'False
         Height          =   345
         Left            =   5100
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   4365
         Width           =   5655
      End
      Begin VB.CommandButton cmdCopyRequest 
         Caption         =   "&Copy Request"
         Enabled         =   0   'False
         Height          =   345
         Left            =   10770
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   4365
         Width           =   5655
      End
      Begin VB.ComboBox cmbmsfgResult 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         ItemData        =   "frmRequest.frx":0000
         Left            =   1590
         List            =   "frmRequest.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3000
         Visible         =   0   'False
         Width           =   2805
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgExamDet 
         Height          =   4395
         Left            =   16995
         TabIndex        =   23
         Top             =   5325
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   7752
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         Enabled         =   0   'False
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin VB.ComboBox cmbSectionFilter 
         Enabled         =   0   'False
         Height          =   330
         Left            =   13110
         Style           =   2  'Dropdown List
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   4875
         Width           =   3315
      End
      Begin VB.CommandButton cmdSaveDemo 
         Caption         =   "&Add"
         Height          =   345
         Left            =   270
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   4365
         Width           =   2265
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2550
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   4365
         Width           =   2265
      End
      Begin VB.TextBox txtmsfgResult 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Left            =   3120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1590
         Visible         =   0   'False
         Width           =   1245
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   3825
         Left            =   270
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   480
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   6747
         _Version        =   393216
         Rows            =   1
         Cols            =   12
         FixedRows       =   0
         GridColorFixed  =   0
         FocusRect       =   0
         HighLight       =   0
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
         _Band(0).Cols   =   12
      End
      Begin VB.TextBox txtmsfgReqDet 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Left            =   3300
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   6900
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.ComboBox cmbmsfgReqDet 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         ItemData        =   "frmRequest.frx":0004
         Left            =   480
         List            =   "frmRequest.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   8160
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.CommandButton cmdSaveReq 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   345
         Left            =   270
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   9825
         Width           =   2265
      End
      Begin VB.CommandButton cmdCancelReq 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   345
         Left            =   2550
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   9825
         Width           =   2265
      End
      Begin VB.TextBox txtmsfgExam 
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   525
         Left            =   5190
         MaxLength       =   30
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   7560
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton cmdSaveExam 
         Caption         =   "&Save Request"
         Enabled         =   0   'False
         Height          =   345
         Left            =   5100
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   9825
         Width           =   3765
      End
      Begin VB.CommandButton cmdDelReq 
         Caption         =   "&Delete Request"
         Enabled         =   0   'False
         Height          =   345
         Left            =   16950
         TabIndex        =   2
         Top             =   4365
         Width           =   3765
      End
      Begin VB.CommandButton cmdDeleteExam 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   345
         Left            =   12660
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   9825
         Width           =   3765
      End
      Begin VB.CommandButton cmdPrintBarCode 
         Caption         =   "&Print Barcode"
         Enabled         =   0   'False
         Height          =   345
         Left            =   8880
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   9825
         Width           =   3765
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgRequest 
         Height          =   3825
         Left            =   5100
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   480
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   6747
         _Version        =   393216
         GridColor       =   -2147483641
         GridColorFixed  =   -2147483641
         GridColorUnpopulated=   16777215
         AllowBigSelection=   0   'False
         Enabled         =   0   'False
         FocusRect       =   0
         GridLinesFixed  =   1
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
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgRequestDet 
         Height          =   4515
         Left            =   270
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   5265
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   7964
         _Version        =   393216
         Rows            =   1
         Cols            =   12
         FixedRows       =   0
         GridColorFixed  =   0
         FocusRect       =   0
         HighLight       =   0
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
         _Band(0).Cols   =   12
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgExam 
         Height          =   4515
         Left            =   5100
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   5265
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   7964
         _Version        =   393216
         GridColor       =   -2147483647
         GridColorFixed  =   -2147483640
         GridColorUnpopulated=   16777215
         FocusRect       =   0
         HighLight       =   2
         GridLinesFixed  =   1
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
         _Band(0).Cols   =   2
      End
      Begin MSComctlLib.ImageList imglstIcon 
         Left            =   4680
         Top             =   4980
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
               Picture         =   "frmRequest.frx":0008
               Key             =   "newrequest"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRequest.frx":0402
               Key             =   "checkin"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRequest.frx":0826
               Key             =   "ongoing"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRequest.frx":0C4B
               Key             =   "ready"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRequest.frx":104F
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRequest.frx":1457
               Key             =   "new"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRequest.frx":1853
               Key             =   "printed"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRequest.frx":1C72
               Key             =   "saved"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRequest.frx":2085
               Key             =   "unapprove"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   12330
         TabIndex        =   21
         Top             =   4965
         Width           =   660
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Demographics"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   330
         TabIndex        =   18
         Top             =   210
         Width           =   1905
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Requests"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   5160
         TabIndex        =   17
         Top             =   210
         Width           =   1845
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Request Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   330
         TabIndex        =   16
         Top             =   4965
         Width           =   1845
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Exam"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   5160
         TabIndex        =   15
         Top             =   4965
         Width           =   1845
      End
   End
End
Attribute VB_Name = "frmRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private intRowCtr As Integer
'Private intColCtr As Integer
'Private rsrequest As New ADODB.Recordset
'Private rsSearch As New ADODB.Recordset
'Private lenTxtmsfgExam As Integer
'Private intBarCodeCtr As Integer
'Private ENA_SCOL_DEMO_NEW As String
'Private ENA_SCOL_DEMO_EDT As String
'Private ENA_SCOL_REQD_NEW As String
'Private ENA_SCOL_REQD_EDT As String
'Private CUR_DEMO_NEW As String
'Private CUR_DEMO_EDT As String
'Private AUTO_COMP_AGE As String
'Private AUTO_CHECK_IN As String
'Private blnCopiedRequest As Boolean
'Private blnWithSettingsDemo As Boolean
'Private blnWithSettingsReqDet As Boolean
'
'
'Private Sub cmbmsfgReqDet_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 38 Then
'        If msfgRequestDet.Row > 1 Then
'            msfgRequestDet_LeaveCell
'            For intRowCtr = msfgRequestDet.Row - 1 To 0 Step -1
'                If msfgRequestDet.TextMatrix(intRowCtr, 2) = "1" And _
'                    msfgRequestDet.TextMatrix(intRowCtr, 3) = "1" Then
'                    msfgRequestDet.Row = intRowCtr
'                    Exit For
'                End If
'            Next
'            msfgRequestDet.SetFocus
'            msfgRequestDet_Click
'        Else
'           cmbmsfgReqDet.SetFocus
'        End If
'    ElseIf KeyCode = 40 Then
'        If msfgRequestDet.Row < msfgRequestDet.Rows - 1 Then
'            msfgRequestDet_LeaveCell
'            For intRowCtr = msfgRequestDet.Row + 1 To msfgRequestDet.Rows - 1
'                If msfgRequestDet.TextMatrix(intRowCtr, 2) = "1" And _
'                        msfgRequestDet.TextMatrix(intRowCtr, 3) = "1" Then
'                    msfgRequestDet.Row = intRowCtr
'                    Exit For
'                End If
'            Next
'            msfgRequestDet.SetFocus
'            msfgRequestDet_Click
'            highLight
'        Else
'            cmbmsfgReqDet.SetFocus
'        End If
'    ElseIf KeyCode = vbKeyEscape Then
'        cmdCancelReq_Click
'    End If
'End Sub
'
'Private Sub cmbmsfgReqDet_KeyPress(KeyAscii As Integer)
'    Dim intCurrentRowCbo As Integer
'    If KeyAscii = 13 Then
'        If msfgRequestDet.Row < msfgRequestDet.Rows - 1 Then
'            msfgRequestDet_LeaveCell
'            intCurrentRowCbo = msfgRequestDet.Row
'            For intRowCtr = msfgRequestDet.Row + 1 To msfgRequestDet.Rows - 1
'                If msfgRequestDet.TextMatrix(intRowCtr, 2) = "1" And _
'                        msfgRequestDet.TextMatrix(intRowCtr, 3) = "1" Then
'                    msfgRequestDet.Row = intRowCtr
'                    Exit For
'                End If
'            Next
'            If msfgRequestDet.Row = intCurrentRowCbo Then
'                cmbmsfgReqDet.Visible = False
'                cmbmsfgReqDet.Enabled = False
'                cmdSaveReq.SetFocus
'            End If
'            msfgRequestDet.SetFocus
'            msfgRequestDet_Click
'            highLight
'        Else
'            cmbmsfgReqDet.SetFocus
'        End If
'    End If
'End Sub
'
'Private Sub cmbSectionFilter_Click()
'    With msfgExam
'        .Col = 0
'        If .Rows > 2 And Len(.TextMatrix(1, 1)) > 0 Then
'            txtmsfgExam.Visible = False
'            For intRowCtr = 1 To .Rows - 1
'                .Row = intRowCtr
'                If UCase(Trim(cmbSectionFilter.Text)) = "ALL" Then
'                    .RowHeight(intRowCtr) = 350
'                    Select Case .TextMatrix(.Row, .Col)
'                        Case ""
'                            Set .CellPicture = imglstIcon.ListImages(6).Picture
'                        Case "0"
'                            Set .CellPicture = imglstIcon.ListImages(6).Picture
'                        Case "1"
'                            Set .CellPicture = imglstIcon.ListImages(2).Picture
'                        Case "2"
'                            Set .CellPicture = imglstIcon.ListImages(3).Picture
'                        Case "3"
'                            Set .CellPicture = imglstIcon.ListImages(4).Picture
'                        Case "4"
'                            Set .CellPicture = imglstIcon.ListImages(8).Picture
'                        Case "5"
'                            Set .CellPicture = imglstIcon.ListImages(7).Picture
'                        Case "6"
'                            Set .CellPicture = imglstIcon.ListImages(9).Picture
'                    End Select
'                ElseIf UCase(Trim(.TextMatrix(intRowCtr, 2))) = UCase(Trim(cmbSectionFilter.Text)) Or _
'                        Len(Trim(.TextMatrix(intRowCtr, 2))) = 0 Then
'                    .RowHeight(intRowCtr) = 350
'                    Select Case .TextMatrix(.Row, .Col)
'                        Case ""
'                            Set .CellPicture = imglstIcon.ListImages(6).Picture
'                        Case "0"
'                            Set .CellPicture = imglstIcon.ListImages(6).Picture
'                        Case "1"
'                            Set .CellPicture = imglstIcon.ListImages(2).Picture
'                        Case "2"
'                            Set .CellPicture = imglstIcon.ListImages(3).Picture
'                        Case "3"
'                            Set .CellPicture = imglstIcon.ListImages(4).Picture
'                        Case "4"
'                            Set .CellPicture = imglstIcon.ListImages(8).Picture
'                        Case "5"
'                            Set .CellPicture = imglstIcon.ListImages(7).Picture
'                        Case "6"
'                            Set .CellPicture = imglstIcon.ListImages(9).Picture
'                    End Select
'                ElseIf UCase(Trim(.TextMatrix(intRowCtr, 2))) <> UCase(Trim(cmbSectionFilter.Text)) And _
'                    UCase(Trim(cmbSectionFilter.Text)) <> "ALL" Then
'                    .RowHeight(intRowCtr) = 0
'                    Set .CellPicture = Nothing
'                End If
'            Next
'        End If
'        .Col = 1
'    End With
'End Sub
'
'Private Sub cmdCancel_Click()
'    If cmdSaveDemo.Caption <> "&Update" Then
'        Unload Me
'    Else
'        txtmsfgResult.Visible = False
'        txtmsfgResult.Enabled = False
'        cmbmsfgResult.Visible = False
'        cmbmsfgResult.Enabled = False
'        cmdSaveDemo.Caption = "&Edit"
'
'        With msfgResult
'            For intRowCtr = 0 To .Rows - 1
'                .Row = intRowCtr
'                .CellBackColor = &H80000011
'                .TextMatrix(intRowCtr, 1) = .TextMatrix(intRowCtr, 11)
'            Next
'            .TopRow = 0
'            .Enabled = False
'        End With
'
'        If msfgRequest.Row > 2 Then
'            msfgRequest.Enabled = True
'        End If
'
'        If msfgRequest.Enabled = True Then
'            msfgRequest_EnterCell
'            cmdNewRequest.Enabled = True
'            msfgRequest.SetFocus
'        Else
'            cmdNewRequest.Enabled = True
'            cmdNewRequest.SetFocus
'        End If
'    End If
'End Sub
'
'Private Sub cmdCancelReq_Click()
'
'    With msfgExamDet
'        .ClearStructure
'        .Rows = 2
'    End With
'
'    msfgExam_LeaveCell
'    With msfgExam
'        Set .DataSource = Nothing
'        .Rows = 2
'        .Enabled = False
'        .BackColor = &H80000005
'        For intColCtr = 0 To .Cols - 1
'            .TextMatrix(1, intColCtr) = ""
'        Next
'        .Col = 0
'        Set .CellPicture = imglstIcon.ListImages(6).Picture
'        .CellAlignment = 4
'    End With
'
'    txtmsfgExam.Text = ""
'    txtmsfgExam.Visible = False
'    txtmsfgExam.Enabled = False
'
'    cmdDeleteExam.Enabled = False
'    cmdPrintBarCode.Enabled = False
'    cmdSaveExam.Enabled = False
'
'    cmbSectionFilter.Enabled = False
'    cmbSectionFilter.ListIndex = 0
'    msfgExam.RowHeight(msfgExam.Rows - 1) = 350
'
'    With msfgRequestDet
'        .Enabled = False
'        For intRowCtr = 0 To .Rows - 1
'            .Row = intRowCtr
'            .TextMatrix(intRowCtr, 1) = ""
'            .CellBackColor = &H80000011
'        Next
'        .TopRow = 0
'    End With
'
'    txtmsfgReqDet.Text = ""
'    txtmsfgReqDet.Visible = False
'    txtmsfgReqDet.Enabled = False
'
'    cmbmsfgReqDet.Visible = False
'    cmbmsfgReqDet.Enabled = False
'
'    For intRowCtr = 0 To cmbmsfgReqDet.ListCount - 1
'        cmbmsfgReqDet.RemoveItem (0)
'    Next
'
'    cmdCancelReq.Cancel = False
'    cmdCancelReq.Enabled = False
'    cmdSaveReq.Enabled = False
'    cmdSaveReq.Caption = "Save"
'
'    cmdNewRequest.Enabled = blnWithSettingsReqDet
'
'    If blnCopiedRequest Then
'        msfgRequest.Rows = msfgRequest.Rows - 1
'        msfgRequest.Row = msfgRequest.Rows - 1
'    End If
'
'    blnCopiedRequest = False
'
'    msfgRequest.Col = 1
'    msfgRequest.ColSel = msfgRequest.Cols - 1
'
'    If Len(msfgRequest.TextMatrix(msfgRequest.Row, 12)) > 0 Then
'        cmdCopyRequest.Enabled = True
'        If msfgRequest.TextMatrix(msfgRequest.Row, 12) < 2 Then
'            cmdDelReq.Enabled = blnWithSettingsReqDet
'        Else
'            cmdDelReq.Enabled = False
'        End If
'    Else
'        cmdDelReq.Enabled = False
'        If msfgRequest.Row > 2 Then
'            msfgRequest.RemoveItem (msfgRequest.Rows - 1)
'        End If
'        msfgRequest.Row = msfgRequest.Rows - 1
'    End If
'
'    cmdCancel.Cancel = True
'    cmdCancel.Enabled = True
'    cmdSaveDemo.Enabled = blnWithSettingsDemo
'
'    If msfgRequest.Rows > 1 And Len(msfgRequest.TextMatrix(msfgRequest.Row, 1)) > 0 Then
'        msfgRequest_EnterCell
'        msfgRequest.Enabled = True
'        msfgRequest.SetFocus
'    Else
'        cmdCancel.SetFocus
'    End If
'
'End Sub
'
'Private Sub cmdCancelReq_GotFocus()
'    msfgRequestDet_LeaveCell
'    txtmsfgReqDet.Visible = False
'    txtmsfgReqDet.Enabled = False
'End Sub
'
'Private Sub cmdCopyRequest_Click()
'    Dim intRowToCopy As Integer
'
'    With msfgRequest
'        msfgRequest_LeaveCell
'        blnCopiedRequest = True
'
'        strSQLStatement = "exec sp_getRequestMasterID_select " & intWorkstation
'        Set rsrequest = myConnection.Execute(strSQLStatement)
'
'        intRowToCopy = .Row
'
'        .Rows = .Rows + 1
'        .Row = .Rows - 1
'        .RowHeight(.Row) = 350
'        .ColSel = .Cols - 1
'
'        .TextMatrix(.Row, 1) = Format(Now, "MM/dd/yyyy Hh:mm")
'
'        For intColCtr = 2 To .Cols - 1
'            If intColCtr = 6 Then
'                .TextMatrix(.Row, intColCtr) = rsrequest.Fields(0).Value
'            ElseIf intColCtr = 8 Then
'                .TextMatrix(.Row, intColCtr) = getValue(2, getUserDet(strRequestBy)) & "|" & strRequestBy
'            ElseIf intColCtr = 19 Then
'                .TextMatrix(.Row, intColCtr) = Mid(.TextMatrix(.Row, 8), 1, InStr(1, .TextMatrix(.Row, 8), "|") - 1)
'            Else
'                .TextMatrix(.Row, intColCtr) = .TextMatrix(intRowToCopy, intColCtr)
'            End If
'        Next
'
'        .TopRow = .Row
'        .Enabled = False
'    End With
'
'        rsrequest.Close
'
'        msfgRequest_EnterCell
'        cmdSaveReq.Caption = "Edit"
'        cmdSaveReq_Click
'        cmdSaveReq.Caption = "Save"
'
'        With msfgExam
'            Set .DataSource = Nothing
'            .Rows = 2
'            For intColCtr = 0 To .Cols - 1
'                .TextMatrix(1, intColCtr) = ""
'            Next
'            .Col = 0
'            Set .CellPicture = imglstIcon.ListImages(6).Picture
'            .CellAlignment = 4
'            .Enabled = False
'        End With
'
'        With msfgExamDet
'            .ClearStructure
'            .Rows = 2
'        End With
'
'        txtmsfgResult.Enabled = False
'        cmdSaveDemo.Enabled = False
'
'        cmdCancel.Cancel = False
'        cmdCancel.Enabled = False
'
'        cmdCancelReq.Enabled = True
'        cmdCancelReq.Cancel = True
'
'        cmdNewRequest.Enabled = False
'        cmdCopyRequest.Enabled = False
'        cmdDelReq.Enabled = False
'
'        cmdSaveExam.Enabled = False
'        cmdPrintBarCode.Enabled = False
'
'End Sub
'
'Private Sub cmdDeleteExam_Click()
''    With msfgExam
''        msfgExam_LeaveCell
''        If Len(.TextMatrix(.Row, 0)) > 0 Then
''            If CInt(.TextMatrix(.Row, 0)) < 2 Then
''                If Len(Trim(.TextMatrix(.Row, 10))) > 0 Then
''                    strSQLStatement = "EXEC sp_tbLabRequestValues_delete '" & .TextMatrix(.Row, 4) & "'" & _
''                                                ",'" & .TextMatrix(.Row, 5) & "'" & _
''                                                ",'" & .TextMatrix(.Row, 1) & "'" & _
''                                                ",'" & strRequestBy & "'"
''                    Set rsSearch = myConnection.Execute(strSQLStatement)
''
''                    'AdditionalWorks "2", .TextMatrix(.Row, 2), .TextMatrix(.Row, 7), .TextMatrix(.Row, 11), strRequestBy
''                End If
''
''                DelExamDet .TextMatrix(.Row, 1)
''
''                If .Rows > 2 Then
''                    .RemoveItem msfgExam.Row
''                Else
''                    For intRowCtr = 0 To .Cols - 1
''                        .TextMatrix(.Row, intRowCtr) = ""
''                    Next
''                End If
''
''
''                If Len(Trim(.TextMatrix(1, 9))) > 0 Then
''                    For intRowCtr = 1 To .Rows - 1
''                        If Len(Trim(.TextMatrix(intRowCtr, 7))) > 0 Then
''                            AdditionalWorks "1", .TextMatrix(intRowCtr, 2), .TextMatrix(intRowCtr, 1), .TextMatrix(intRowCtr, 7), .TextMatrix(intRowCtr, 11), strRequestBy
''                        End If
''                    Next
''                End If
''
''                .SetFocus
''                msfgExam_EnterCell
''            End If
''        Else
''            txtmsfgExam.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
''            txtmsfgExam.Visible = True
''            txtmsfgExam.Enabled = True
''            txtmsfgExam.SetFocus
''        End If
''    End With
'End Sub
'
'Private Sub cmdDelReq_Click()
''    With msfgRequest
''        If Len(.TextMatrix(.Row, 0)) > 0 Then
''            If .ColSel > 1 Then
''                If MsgBox("Are you sure you want to delete this request", vbYesNo + vbQuestion, "Request") = vbYes Then
''                    strSQLStatement = "EXEC sp_tbLabRequestMaster_delete '" & .TextMatrix(.Row, 6) & "'" & _
''                                                ",'" & .TextMatrix(.Row, 7) & "','" & strRequestBy & "'"
''                    Set rsSearch = myConnection.Execute(strSQLStatement)
''
''                    strSQLStatement = "sp_tbLabRequestValues_delete '" & .TextMatrix(.Row, 6) & "',NULL,NULL,'" & strRequestBy & "'"
''                    Set rsSearch = myConnection.Execute(strSQLStatement)
''
''                    If .Rows > 2 Then
''                        .RemoveItem .Row
''                    Else
''                        For intRowCtr = 0 To .Cols - 1
''                            .TextMatrix(.Row, intRowCtr) = ""
''                        Next
''                    End If
''
''
''                    .SetFocus
''                    msfgRequest_EnterCell
''                Else
''                    .SetFocus
''                End If
''            End If
''        Else
''            .SetFocus
''        End If
''    End With
''
'End Sub
'
'Private Sub cmdNewRequest_Click()
'
'    With msfgRequest
'        .Enabled = False
'        If Len(msfgRequest.TextMatrix(msfgRequest.Rows - 1, 0)) > 0 Then
'            .Rows = .Rows + 1
'        End If
'
'        .Row = .Rows - 1
'        .RowHeight(.Row) = 330
'        .TopRow = .Row
'        .CellAlignment = 0
'
'        For intRowCtr = 1 To msfgRequest.Cols - 1
'            .Col = intRowCtr
'            .CellBackColor = &H8000000D
'            .CellForeColor = vbWhite
'        Next
'
'        .Enabled = False
'    End With
'
'    With msfgRequestDet
'        strSQLStatement = "exec sp_getRequestMasterID_select " & intWorkstation
'        Set rsrequest = myConnection.Execute(strSQLStatement)
''0=requestdate
''1=ward
''2=room
''3=patienttype
''4=location
''5=requestid
''6=requestpid
''7=requestby
''8=age
''9=priority
''10=communication
''11=status
''12=physician
''13=reserve1
''14=reserve2
''15=reserve3
''16=reserve4
''17=reserve5
'        With msfgRequestDet
'            For intRowCtr = 0 To .Rows - 1
'                .TextMatrix(intRowCtr, 1) = ""
'                .TextMatrix(intRowCtr, 11) = ""
'                .Row = intRowCtr
'                If .TextMatrix(.Row, 2) = "1" Then
'                    .CellBackColor = &H80000014
'                End If
'            Next
'        End With
'
'        .Row = 0
'        .TextMatrix(0, 1) = Format(Now, "MM/dd/yyyy Hh:mm")
'        .TextMatrix(5, 1) = rsrequest.Fields(0).Value
'        .TextMatrix(6, 1) = PatientID
'        .TextMatrix(7, 1) = strRequestByName
'        .TextMatrix(7, 10) = strRequestBy
'        .TextMatrix(11, 1) = "0"
'
'        If Len(Trim(msfgResult.TextMatrix(5, 1))) > 0 And AUTO_COMP_AGE = "1" And _
'                    IsDate(msfgResult.TextMatrix(5, 1)) Then
'            .TextMatrix(8, 1) = CompAge(Now, msfgResult.TextMatrix(5, 1))
'        End If
'
'        If ENA_SCOL_REQD_NEW = "1" Then
'            For intRowCtr = 0 To .Rows - 1
'                If Len(Trim(.TextMatrix(intRowCtr, 7))) > 0 Then
'                    .Row = intRowCtr
'                    Exit For
'                End If
'            Next
'        Else
'            For intRowCtr = 0 To .Rows - 1
'                If Len(Trim(.TextMatrix(intRowCtr, 1))) = 0 And _
'                            Trim(.TextMatrix(intRowCtr, 2)) = "1" And _
'                            Trim(.TextMatrix(intRowCtr, 3)) = "1" Then
'                    .Row = intRowCtr
'                    Exit For
'                End If
'            Next
'        End If
'
'        If UCase(.TextMatrix(.Row, 5)) = "TEXT" Then
'            txtmsfgReqDet.Move (.CellLeft + .Left), (.CellTop + .Top), .CellWidth, .CellHeight
'            txtmsfgReqDet.Visible = True
'            txtmsfgReqDet.Enabled = True
'            txtmsfgReqDet.Text = .TextMatrix(.Row, 1)
'            txtmsfgReqDet.SetFocus
'        ElseIf UCase(.TextMatrix(.Row, 5)) = "COMBO" Then
'            cmbmsfgReqDet.Visible = True
'            cmbmsfgReqDet.Enabled = True
'            cmbmsfgReqDet.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
'            setListcmbmsfgReqDet .TextMatrix(.Row, 6)
'            If Len(.TextMatrix(.Row, .Col)) > 0 Then
'                For intRowCtr = 0 To cmbmsfgReqDet.ListCount - 1
'                    If .TextMatrix(.Row, 1) = cmbmsfgReqDet.List(intRowCtr) Then
'                        cmbmsfgReqDet.ListIndex = intRowCtr
'                    End If
'                Next
'            End If
'            cmbmsfgReqDet.SetFocus
'        End If
'
'
'        .Enabled = True
'    End With
'
'    msfgResult.Enabled = False
'
'    txtmsfgResult.Enabled = False
'    cmdSaveDemo.Enabled = False
'
'    cmdSaveReq.Enabled = True
'    cmdSaveReq.Caption = "Save"
'
'    cmdCancel.Cancel = False
'    cmdCancel.Enabled = False
'
'    cmdCancelReq.Enabled = True
'    cmdCancelReq.Cancel = True
'
'    cmdNewRequest.Enabled = False
'    cmdCopyRequest.Enabled = False
'    cmdDelReq.Enabled = False
'
'
'    With msfgExam
'        Set .DataSource = Nothing
'        .Rows = 2
'        For intColCtr = 0 To .Cols - 1
'            .TextMatrix(1, intColCtr) = ""
'        Next
'        .Col = 0
'        Set .CellPicture = imglstIcon.ListImages(6).Picture
'        .CellAlignment = 4
'    End With
'
'    With msfgExamDet
'        .ClearStructure
'        .Rows = 2
'    End With
'
'End Sub
'
'Private Sub cmdPrintBarCode_Click()
'    If msfgExam.ColSel > 0 Then
'        printBarCode msfgExam.TextMatrix(msfgExam.Row, 6)
'        msfgExam.SetFocus
'    End If
'End Sub
'
'Private Sub cmdSaveDemo_Click()
'    Dim rsSave As New ADODB.Recordset
'
'    With msfgResult
'        Select Case cmdSaveDemo.Caption
'            Case "&Add"
'                msfgResult_LeaveCell
'                If Len(.TextMatrix(2, 1)) > 0 And Len(.TextMatrix(3, 1)) > 0 And (Len(Trim(.TextMatrix(5, 1))) = 0 Or IsDate(Trim(.TextMatrix(5, 1)))) Then
'
'                    strSQLStatement = "EXEC sp_getPatientID_select"
'                    Set rsSearch = myConnection.Execute(strSQLStatement)
'                    PatientID = rsSearch.Fields("PID")
'
'
'                    strSQLStatement = "exec sp_PatientDemo_insert '" & Trim(PatientID) & "'"
'
'                    For intRowCtr = 1 To .Rows - 1
'                        strSQLStatement = strSQLStatement & ",'" & _
'                                            IIf(InStr(1, .TextMatrix(intRowCtr, 1), "'") > 0, _
'                                            Replace(.TextMatrix(intRowCtr, 1), "'", "''"), _
'                                            .TextMatrix(intRowCtr, 1)) & "'"
'                    Next
'
'                    strSQLStatement = strSQLStatement & ",'" & strRequestBy & "'"
'
'                    Set rsSave = myConnection.Execute(strSQLStatement)
'                    .Enabled = False
'
'                    For intRowCtr = 0 To .Rows - 1
'                        .Row = intRowCtr
'                        .CellBackColor = &H80000011
'                    Next
'
'                    For intRowCtr = 0 To .Rows - 1
'                        .TextMatrix(intRowCtr, 11) = .TextMatrix(intRowCtr, 1)
'                    Next
'
'                    txtmsfgResult.Visible = False
'                    txtmsfgResult.Enabled = False
'                    cmbmsfgResult.Visible = False
'                    cmbmsfgResult.Enabled = False
'
'                    cmdSaveDemo.Caption = "&Edit"
'                    cmdNewRequest.Enabled = True
'                    cmdNewRequest.SetFocus
'                Else
'                    If .TextMatrix(.Row, 6) = "TEXT" Then
'                        txtmsfgResult.SetFocus
'                    Else
'                        cmbmsfgResult.SetFocus
'                    End If
'                End If
'            Case "&Update"
'                msfgResult_LeaveCell
'                If Len(.TextMatrix(2, 1)) > 0 And Len(.TextMatrix(3, 1)) > 0 And (Len(Trim(.TextMatrix(5, 1))) = 0 Or IsDate(Trim(.TextMatrix(5, 1)))) Then
'
'                    strSQLStatement = "exec sp_PatientDemo_update '" & Trim(PatientID) & "'"
'
'                    For intRowCtr = 1 To .Rows - 1
'                        strSQLStatement = strSQLStatement & ",'" & _
'                                            IIf(InStr(1, .TextMatrix(intRowCtr, 1), "'") > 0, _
'                                            Replace(.TextMatrix(intRowCtr, 1), "'", "''"), _
'                                            .TextMatrix(intRowCtr, 1)) & "'"
'                    Next
'
'                    strSQLStatement = strSQLStatement & ",'" & strRequestBy & "'"
'
'                    Set rsSave = myConnection.Execute(strSQLStatement)
'
'                    cmdNewRequest.Enabled = True
'                    cmdNewRequest.SetFocus
'
'                    .Enabled = False
'                    For intRowCtr = 0 To .Rows - 1
'                        .Row = intRowCtr
'                        .CellBackColor = &H80000011
'                        .TextMatrix(.Row, 11) = .TextMatrix(.Row, 1)
'                    Next
'
'                    txtmsfgResult.Visible = False
'                    txtmsfgResult.Enabled = False
'                    cmbmsfgResult.Visible = False
'                    cmbmsfgResult.Enabled = False
'
'                    cmdSaveDemo.Caption = "&Edit"
'                    cmdSaveReq.Enabled = True
'
'                Else
'                    If .TextMatrix(.Row, 6) = "TEXT" Then
'                        If txtmsfgResult.Enabled Then
'                            txtmsfgResult.SetFocus
'                        Else
'                            cmdSaveDemo.SetFocus
'                        End If
'                    Else
'                        If txtmsfgResult.Enabled Then
'                            cmbmsfgResult.SetFocus
'                        Else
'                            cmdSaveDemo.SetFocus
'                        End If
'                    End If
'                End If
'
'            Case "&Edit"
'                msfgRequest.Enabled = False
'                cmdSaveReq.Enabled = False
'                cmdNewRequest.Enabled = False
'                cmdCopyRequest.Enabled = False
'                cmdDelReq.Enabled = False
'                .Enabled = True
'
'                For intRowCtr = 0 To .Rows - 1
'                    .Row = intRowCtr
'                    If .TextMatrix(.Row, 3) = "1" And .TextMatrix(.Row, 4) = "1" Then
'                        .CellBackColor = &H80000005
'                    End If
'                Next
'
'                For intRowCtr = 0 To .Rows - 1
'                    If .TextMatrix(intRowCtr, 9) = "1" Then
'                        .Row = intRowCtr
'                        Exit For
'                    End If
'                Next
'
'                msfgResult_Click
'                cmdSaveDemo.Caption = "&Update"
'        End Select
'    End With
'
'End Sub
'
'Private Sub cmdSaveExam_Click()
'    Dim strBarCodeToPrint As String
'    Dim strBarCodePrinted As String
'    msfgExam_LeaveCell
'    With msfgExam
'        For intRowCtr = 1 To .Rows - 1
'            If Len(Trim(.TextMatrix(intRowCtr, 6))) = 0 And Len(Trim(.TextMatrix(intRowCtr, 9))) = 0 And _
'               Len(Trim(.TextMatrix(intRowCtr, 8))) > 0 And Len(Trim(.TextMatrix(intRowCtr, 1))) > 0 Then
'                If AUTO_CHECK_IN = "1" Then
'                    strSQLStatement = "EXEC sp_getRVSampleID_select '" & .TextMatrix(intRowCtr, 4) & "'," & _
'                                        "'" & msfgRequest.TextMatrix(msfgRequest.Row, 1) & "'," & _
'                                        "'" & .TextMatrix(intRowCtr, 3) & "','" & .TextMatrix(intRowCtr, 7) & "'"
'
'                    Set rsrequest = myConnection.Execute(strSQLStatement)
'
'                    .TextMatrix(intRowCtr, 0) = "1"
'                    .TextMatrix(intRowCtr, 5) = "1"
'                    .TextMatrix(intRowCtr, 6) = RTrim(rsrequest.Fields(0).Value)
'
'                    strSQLStatement = "EXEC sp_tbLabRequestValues_insert " & _
'                                    "'" & .TextMatrix(intRowCtr, 3) & "',NULL," & _
'                                    "'" & .TextMatrix(intRowCtr, 4) & "'," & _
'                                    "'" & .TextMatrix(intRowCtr, 1) & "'," & _
'                                    "'" & .TextMatrix(intRowCtr, 5) & "'," & _
'                                    "'" & .TextMatrix(intRowCtr, 6) & "'," & _
'                                    "'" & strRequestBy & "'"
'
'                    Set rsrequest = myConnection.Execute(strSQLStatement)
'
'                    .Row = intRowCtr
'                    .Col = 0
'                    Set .CellPicture = imglstIcon.ListImages(2).Picture
'                    .Row = .Rows - 1
'                    .Col = 1
'                Else
'                    .TextMatrix(intRowCtr, 0) = "0"
'                    .TextMatrix(intRowCtr, 5) = "0"
'                    .TextMatrix(intRowCtr, 6) = ""
'
'                    strSQLStatement = "EXEC sp_tbLabRequestValues_insert " & _
'                                    "'" & .TextMatrix(intRowCtr, 3) & "',NULL," & _
'                                    "'" & .TextMatrix(intRowCtr, 4) & "'," & _
'                                    "'" & .TextMatrix(intRowCtr, 1) & "'," & _
'                                    "'" & .TextMatrix(intRowCtr, 5) & "'," & _
'                                    "'" & .TextMatrix(intRowCtr, 6) & "'," & _
'                                    "'" & strRequestBy & "'"
'
'                    Set rsrequest = myConnection.Execute(strSQLStatement)
'
'                End If
'            End If
'        Next
'
'        If AUTO_CHECK_IN = "1" Then
'            printBarCode "", msfgRequest.TextMatrix(msfgRequest.Row, 6)
'        End If
'
'        rPatientID = msfgResult.TextMatrix(0, 1)
'        rRequestID = msfgRequest.TextMatrix(msfgRequest.Row, 6)
'
'        If showSummaryRequest(msfgRequest.TextMatrix(msfgRequest.Row, 6)) = "1" Then
'            Unload Me
'            Unload frmSearchRequest
'        Else
'            If Not isExiting Then
'                With msfgExam
'                    .Row = .Rows - 1
'                    .Col = 1
'                    txtmsfgExam.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
'                    txtmsfgExam.Visible = True
'                    txtmsfgExam.Enabled = True
'                    txtmsfgExam.SetFocus
'                End With
'            End If
'        End If
'
'    End With
'End Sub
'
'Private Sub cmdSaveReq_Click()
'    Dim strMandatory As String
'    Dim rsRID As New ADODB.Recordset
'
'    If cmdSaveReq.Caption = "Edit" Then
'        With msfgRequestDet
'            For intRowCtr = 0 To .Rows - 1
'                .Row = intRowCtr
'                If .TextMatrix(.Row, 2) = "1" Then
'                    .CellBackColor = &H80000014
'                End If
'            Next
'
'            If ENA_SCOL_REQD_EDT = "1" Then
'                For intRowCtr = 0 To .Rows - 1
'                    If Len(Trim(.TextMatrix(intRowCtr, 8))) > 0 Then
'                        .Row = intRowCtr
'                        Exit For
'                    End If
'                Next
'            Else
'                For intRowCtr = 0 To .Rows - 1
'                    If Len(Trim(.TextMatrix(intRowCtr, 1))) = 0 And _
'                                Trim(.TextMatrix(intRowCtr, 2)) = "1" And _
'                                Trim(.TextMatrix(intRowCtr, 3)) = "1" Then
'                        .Row = intRowCtr
'                        Exit For
'                    End If
'                Next
'            End If
'            msfgRequestDet_Click
'        End With
'
'        msfgRequestDet.Enabled = True
'
'        cmbSectionFilter.Enabled = True
'
'        msfgExam.BackColor = &H80000014
'        msfgExam.Enabled = True
'
'        cmdSaveReq.Enabled = True
'        cmdSaveReq.Caption = "&Update"
'
'        cmdPrintBarCode.Enabled = True
'        cmdSaveExam.Enabled = True
'
'        cmdSaveDemo.Enabled = False
'        cmdCancel.Enabled = False
'
'        cmdNewRequest.Enabled = False
'        cmdCopyRequest.Enabled = False
'        cmdDelReq.Enabled = False
'        msfgRequest.Enabled = False
'        cmdCancelReq.Enabled = True
'
'    Else
'        msfgRequestDet_LeaveCell
'
'        If Len(msfgRequestDet.TextMatrix(12, 1)) > 0 And Len(msfgRequestDet.TextMatrix(12, 10)) = 0 Then
'            txtmsfgReqDet.SetFocus
'            Exit Sub
'        End If
'
'        With msfgRequestDet
'            For intRowCtr = 0 To .Rows - 1
'                If Len(Trim(.TextMatrix(intRowCtr, 1))) = 0 And Trim(.TextMatrix(intRowCtr, 3)) = "1" And _
'                    Trim(.TextMatrix(intRowCtr, 4) = "1") Then
'                    strMandatory = strMandatory & IIf(Len(strMandatory) = 0, .TextMatrix(intRowCtr, 0), ", " & .TextMatrix(intRowCtr, 0))
'                End If
'            Next
'
'            If Len(strMandatory) > 0 Then
'                If InStr(1, strMandatory, ",") > 0 Then
'                    MsgBox strMandatory & " Are Required Fields", vbInformation + vbOKOnly, "Request"
'                Else
'                    MsgBox strMandatory & " Is Required", vbInformation + vbOKOnly, "Request"
'                End If
'
'                msfgRequestDet_Click
'
'                Exit Sub
'            End If
'
'            txtmsfgReqDet.Visible = False
'            txtmsfgReqDet.Enabled = False
'            cmbmsfgReqDet.Visible = False
'            cmbmsfgReqDet.Enabled = False
'    '0=requestdate
'    '1=ward
'    '2=room
'    '3=patienttype
'    '4=location
'    '5=requestid
'    '6=requestpid
'    '7=requestby
'    '8=age
'    '9=priority
'    '10=comment
'    '11=status
'    '12=physician
'    '13=reserve1
'    '14=reserve2
'    '15=reserve3
'    '16=reserve4
'    '17=reserve5
'            If cmdSaveReq.Caption = "Save" Then
'                Do
'                    strSQLStatement = "EXEC sp_CheckRequestID_Select '" & .TextMatrix(5, 1) & "'"
'                    Set rsrequest = myConnection.Execute(strSQLStatement)
'
'                    If rsrequest.RecordCount > 0 Then
'                        strSQLStatement = "exec sp_getRequestMasterID_select " & intWorkstation
'                        Set rsrequest = myConnection.Execute(strSQLStatement)
'                        .TextMatrix(5, 1) = RTrim(rsrequest.Fields(0))
'                    End If
'
'                Loop Until rsrequest.RecordCount = 0
'
'                strSQLStatement = "exec sp_tbLabRequestMaster_insert '" & .TextMatrix(1, 1) & "'," & _
'                                            "'" & .TextMatrix(2, 1) & "'," & _
'                                            "'" & .TextMatrix(3, 1) & "'," & _
'                                            "'" & .TextMatrix(4, 1) & "'," & _
'                                            "'" & .TextMatrix(5, 1) & "'," & _
'                                            "'" & .TextMatrix(6, 1) & "'," & _
'                                            "'" & strRequestBy & "'," & _
'                                            "'" & .TextMatrix(8, 1) & "'," & _
'                                            "'" & .TextMatrix(9, 1) & "'," & _
'                                            "'" & .TextMatrix(10, 1) & "'," & _
'                                            "'" & .TextMatrix(11, 1) & "'," & _
'                                            "'" & .TextMatrix(12, 10) & "'," & _
'                                            "'" & .TextMatrix(13, 1) & "'," & _
'                                            "'" & .TextMatrix(14, 1) & "'," & _
'                                            "'" & .TextMatrix(15, 1) & "'," & _
'                                            "'" & .TextMatrix(16, 1) & "'," & _
'                                            "'" & .TextMatrix(17, 1) & "'"
'                cmdSaveReq.Caption = "&Update"
'                blnCopiedRequest = False
'            ElseIf cmdSaveReq.Caption = "&Update" Then
'                strSQLStatement = "exec sp_tbLabRequestMaster_insert '" & .TextMatrix(1, 1) & "'," & _
'                                            "'" & .TextMatrix(2, 1) & "'," & _
'                                            "'" & .TextMatrix(3, 1) & "'," & _
'                                            "'" & .TextMatrix(4, 1) & "'," & _
'                                            "'" & .TextMatrix(5, 1) & "'," & _
'                                            "'" & .TextMatrix(6, 1) & "'," & _
'                                            "'" & strRequestBy & "'," & _
'                                            "'" & .TextMatrix(8, 1) & "'," & _
'                                            "'" & .TextMatrix(9, 1) & "'," & _
'                                            "'" & .TextMatrix(10, 1) & "'," & _
'                                            "'" & .TextMatrix(11, 1) & "'," & _
'                                            "'" & .TextMatrix(12, 10) & "'," & _
'                                            "'" & .TextMatrix(13, 1) & "'," & _
'                                            "'" & .TextMatrix(14, 1) & "'," & _
'                                            "'" & .TextMatrix(15, 1) & "'," & _
'                                            "'" & .TextMatrix(16, 1) & "'," & _
'                                            "'" & .TextMatrix(17, 1) & "'"
'
'            End If
'
'            Set rsrequest = myConnection.Execute(strSQLStatement)
'
'        End With
'
'        With msfgRequest
'            '.TextMatrix(.Row, 0) = msfgRequestDet.TextMatrix(11, 1)
'            .TextMatrix(.Row, 1) = msfgRequestDet.TextMatrix(0, 1)
'            .TextMatrix(.Row, 2) = msfgRequestDet.TextMatrix(1, 1)
'            .TextMatrix(.Row, 3) = msfgRequestDet.TextMatrix(2, 1)
'            .TextMatrix(.Row, 4) = msfgRequestDet.TextMatrix(3, 1)
'            .TextMatrix(.Row, 5) = msfgRequestDet.TextMatrix(4, 1)
'            .TextMatrix(.Row, 6) = msfgRequestDet.TextMatrix(5, 1)
'            .TextMatrix(.Row, 7) = msfgRequestDet.TextMatrix(6, 1)
'            .TextMatrix(.Row, 8) = msfgRequestDet.TextMatrix(7, 1) & "|" & strRequestBy
'            .TextMatrix(.Row, 9) = msfgRequestDet.TextMatrix(8, 1)
'            .TextMatrix(.Row, 10) = msfgRequestDet.TextMatrix(9, 1)
'            .TextMatrix(.Row, 11) = msfgRequestDet.TextMatrix(10, 1)
'            .TextMatrix(.Row, 12) = msfgRequestDet.TextMatrix(11, 1)
'            .TextMatrix(.Row, 13) = msfgRequestDet.TextMatrix(12, 1) & "|" & msfgRequestDet.TextMatrix(12, 10)
'            .TextMatrix(.Row, 14) = msfgRequestDet.TextMatrix(13, 1)
'            .TextMatrix(.Row, 15) = msfgRequestDet.TextMatrix(14, 1)
'            .TextMatrix(.Row, 16) = msfgRequestDet.TextMatrix(15, 1)
'            .TextMatrix(.Row, 17) = msfgRequestDet.TextMatrix(16, 1)
'            .TextMatrix(.Row, 18) = msfgRequestDet.TextMatrix(17, 1)
'            .TextMatrix(.Row, 19) = getValue(2, getUserDet(strRequestBy))
'        End With
'
'        cmbSectionFilter.Enabled = True
'        cmdSaveExam.Enabled = True
'        cmdPrintBarCode.Enabled = True
'
'        With msfgExam
'            .Enabled = True
'            .Row = .Rows - 1
'            .BackColor = &H80000014
'            .Col = 1
'            .CellAlignment = 0
'            txtmsfgExam.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
'            txtmsfgExam.Enabled = True
'            txtmsfgExam.Visible = True
'            txtmsfgExam.SetFocus
'        End With
'    End If
'End Sub
'
'Private Sub cmdSaveReq_GotFocus()
'    txtmsfgReqDet.TabStop = False
'    msfgRequestDet.TabStop = False
'End Sub
'
'
'
'Private Sub Form_Load()
'    setFormColors Me
'
'    msfgRequest.BackColorFixed = GetSetting("LMS", "FormColors", "TitleBar", &HA66B2B)
'    msfgRequest.ForeColorFixed = GetSetting("LMS", "FormColors", "TitleText", &HFFFFFF)
'
'    msfgRequestDet.BackColorFixed = GetSetting("LMS", "FormColors", "TitleBar", &HA66B2B)
'    msfgRequestDet.ForeColorFixed = GetSetting("LMS", "FormColors", "TitleText", &HFFFFFF)
'
'    msfgExam.BackColorFixed = GetSetting("LMS", "FormColors", "TitleBar", &HA66B2B)
'    msfgExam.ForeColorFixed = GetSetting("LMS", "FormColors", "TitleText", &HFFFFFF)
'
'    If searchObj Is Nothing Then
'        Set searchObj = CreateObject("SearchCodes.ClassSearch")
'    End If
'
'    strSQLStatement = "SELECT SECTIONNAME FROM tbLabSection"
'
'    Set rsrequest = myConnection.Execute(strSQLStatement)
'
'    cmbSectionFilter.AddItem ("ALL")
'    If rsrequest.RecordCount > 0 Then
'        For intRowCtr = 0 To rsrequest.RecordCount - 1
'            cmbSectionFilter.AddItem (rsrequest.Fields(0).Value)
'            rsrequest.MoveNext
'        Next
'    End If
'    cmbSectionFilter.ListIndex = 0
'
'    ENA_SCOL_DEMO_NEW = getGlobalSettings("ENA_SCOL_DEMO_NEW", "RE_DEMO", "0")
'    CUR_DEMO_NEW = getGlobalSettings("CUR_DEMO_NEW", "RE_DEMO", "")
'
'    ENA_SCOL_DEMO_EDT = getGlobalSettings("ENA_SCOL_DEMO_EDT", "RE_DEMO", "0")
'    CUR_DEMO_EDT = getGlobalSettings("CUR_DEMO_EDT", "RE_DEMO", "")
'
'    ENA_SCOL_REQD_NEW = getGlobalSettings("ENA_SCOL_REQD_NEW", "RE_REQ", "0")
'    ENA_SCOL_REQD_EDT = getGlobalSettings("ENA_SCOL_REQD_EDT", "RE_REQ", "0")
'    AUTO_COMP_AGE = getGlobalSettings("AUTO_COMP_AGE", "RE_REQ", "0")
'    AUTO_CHECK_IN = getGlobalSettings("AUTO_CHECK_IN", "RE_REQ", "0")
'
'    If blnNewPatient Then
'        setInitialmsfgResult
'        setInitialmsfgRequest
'        setInitialmsfgRequestDet
'        setmsfgExam
'        msfgRequest.ColSel = msfgRequest.Cols - 1
'        If ENA_SCOL_DEMO_NEW = 1 Then
'            For intRowCtr = 0 To msfgResult.Rows - 1
'                If msfgResult.TextMatrix(intRowCtr, 8) = "1" Then
'                    msfgResult.Row = intRowCtr
'                    Exit For
'                End If
'            Next
'        Else
'            For intRowCtr = 0 To msfgResult.Rows - 1
'                If msfgResult.TextMatrix(intRowCtr, 3) = "1" And msfgResult.TextMatrix(intRowCtr, 4) = "1" Then
'                    msfgResult.Row = intRowCtr
'                    Exit For
'                End If
'            Next
'        End If
'
'        With msfgResult
'            If .TextMatrix(.Row, 3) = "1" And .TextMatrix(.Row, 4) = "1" Then
'                .CellAlignment = 0
'                 If .TextMatrix(.Row, 6) = "TEXT" Then
'                     txtmsfgResult.Move (.CellLeft + .Left), (.CellTop + .Top), .CellWidth, .CellHeight
'                     txtmsfgResult.Visible = True
'                     txtmsfgResult.Enabled = True
'                     txtmsfgResult.Text = .TextMatrix(.Row, .Col)
'                     txtmsfgResult.MaxLength = setLenTxtBox(msfgResult.Name, .TextMatrix(.Row, 3))
'                     txtmsfgResult.TabIndex = 0
'                     txtmsfgResult.TabStop = True
'                 Else
'                     cmbmsfgResult.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
'                     cmbmsfgResult.Visible = True
'                     cmbmsfgResult.Enabled = True
'                     setListcmbmsfgReqDet .TextMatrix(.Row, 6)
'                     If Len(.TextMatrix(.Row, .Col)) > 0 Then
'                         For intRowCtr = 0 To cmbmsfgResult.ListCount - 1
'                             If .TextMatrix(.Row, 1) = cmbmsfgResult.List(intRowCtr) Then
'                                 cmbmsfgResult.ListIndex = intRowCtr
'                             End If
'                         Next
'                     End If
'                     cmbmsfgResult.TabIndex = 0
'                     cmbmsfgResult.TabStop = True
'                 End If
'            End If
'        End With
'
'        cmdSaveDemo.Enabled = blnWithSettingsDemo
'    Else
'        setInitialmsfgRequest
'        setInitialmsfgRequestDet
'        setInitialmsfgResult
'        msfgResult.Enabled = False
'        txtmsfgResult.Enabled = False
'        txtmsfgResult.Visible = False
'        cmdSaveDemo.Caption = "&Edit"
'        cmdSaveDemo.Enabled = False
'        cmdNewRequest.Enabled = False
'        cmdCancel.Cancel = False
'        cmdCancel.Enabled = False
'        cmdCancelReq.Cancel = True
'        cmdCancelReq.Enabled = True
'
'        With msfgRequest
'            If Len(Trim(RequestID)) = 0 Then
'                strSQLStatement = "exec sp_getRequestMasterID_select " & intWorkstation
'                Set rsrequest = myConnection.Execute(strSQLStatement)
'
'                If .Rows > 2 Then
'                    .Rows = .Rows + 1
'                End If
'
'                .Row = .Rows - 1
'                .RowHeight(.Row) = 330
'
'                If ENA_SCOL_REQD_NEW = 1 Then
'                    For intRowCtr = 0 To msfgRequestDet.Rows - 1
'                        msfgRequestDet.Row = intRowCtr
'                        If Len(msfgRequestDet.TextMatrix(msfgRequestDet.Row, 7)) > 0 Then
'                            Exit For
'                        End If
'                    Next
'                Else
'                    For intRowCtr = 0 To msfgRequestDet.Rows - 1
'                        If Len(Trim(msfgRequestDet.TextMatrix(intRowCtr, 1))) = 0 And _
'                            msfgRequestDet.TextMatrix(intRowCtr, 2) = "1" And _
'                            msfgRequestDet.TextMatrix(intRowCtr, 3) = "1" Then
'                            msfgRequestDet.Row = intRowCtr
'                            Exit For
'                        End If
'                    Next
'                End If
'
'                .TopRow = .Rows - 1
'
'                msfgRequestDet.TextMatrix(0, 1) = Format(Now, "MM/dd/yyyy Hh:mm")
'                msfgRequestDet.TextMatrix(5, 1) = RTrim(rsrequest.Fields(0).Value)
'                msfgRequestDet.TextMatrix(6, 1) = RTrim(PatientID)
'                msfgRequestDet.TextMatrix(7, 1) = strRequestByName
'                msfgRequestDet.TextMatrix(7, 10) = strRequestBy
'
'                If Len(Trim(msfgResult.TextMatrix(5, 1))) > 0 And AUTO_COMP_AGE = "1" And _
'                    IsDate(msfgResult.TextMatrix(5, 1)) Then
'                    msfgRequestDet.TextMatrix(8, 1) = CompAge(Now, msfgResult.TextMatrix(5, 1))
'                End If
'                msfgRequestDet.TextMatrix(11, 1) = "0"
'
'                cmdSaveReq.Enabled = blnWithSettingsReqDet
'                cmdSaveReq.Caption = "Save"
'
'                setmsfgExam
'            Else
'                cmbSectionFilter.Enabled = True
'
'                For intRowCtr = 1 To .Rows - 1
'                    .Row = intRowCtr
'                    If .TextMatrix(msfgRequest.Row, 6) = RequestID Then
'                        Exit For
'                    End If
'                Next
'
'                .TopRow = .Row
'
'                For intColCtr = 1 To msfgRequestDet.Rows - 1
'                    msfgRequestDet.Row = intColCtr - 1
'                    msfgRequestDet.Col = 1
'                    msfgRequestDet.CellAlignment = 0
'                    msfgRequestDet.TextMatrix(msfgRequestDet.Row, 1) = .TextMatrix(.Row, intColCtr)
'                Next
'
'                If InStr(1, msfgRequestDet.TextMatrix(7, 1), "|") > 0 Then
'                    msfgRequestDet.TextMatrix(7, 10) = Mid(msfgRequestDet.TextMatrix(7, 1), _
'                                                    InStr(1, msfgRequestDet.TextMatrix(7, 1), "|") + 1)
'                    msfgRequestDet.TextMatrix(7, 1) = Mid(msfgRequestDet.TextMatrix(7, 1), 1, _
'                                                    InStr(1, msfgRequestDet.TextMatrix(7, 1), "|") - 1)
'                End If
'
'                If InStr(1, msfgRequestDet.TextMatrix(12, 1), "|") > 0 Then
'                    msfgRequestDet.TextMatrix(12, 10) = Mid(msfgRequestDet.TextMatrix(12, 1), _
'                                                    InStr(1, msfgRequestDet.TextMatrix(12, 1), "|") + 1)
'                    msfgRequestDet.TextMatrix(12, 1) = Mid(msfgRequestDet.TextMatrix(12, 1), 1, _
'                                                    InStr(1, msfgRequestDet.TextMatrix(12, 1), "|") - 1)
'                End If
'
'
'                If msfgRequest.TextMatrix(msfgRequest.Row, 0) > 1 Then
'                    For intRowCtr = 0 To msfgRequestDet.Rows - 1
'                        msfgRequestDet.Row = intRowCtr
'                        msfgRequestDet.CellBackColor = &H80000011
'                    Next
'                End If
'
'                cmdSaveReq.Enabled = blnWithSettingsReqDet
'                cmdSaveReq.Caption = "Update"
'
'                setmsfgExam RequestID
'                msfgExam.Enabled = blnWithSettingsReqDet
'
'                If msfgRequest.TextMatrix(msfgRequest.Row, 0) > 1 Then
'                    With msfgExam
'                        .BackColor = &H80000011
'                        If .Rows > 2 Then
'                            .RemoveItem (.Rows - 1)
'                        End If
'                    End With
'                End If
'
'                cmdSaveExam.Enabled = blnWithSettingsReqDet
'                cmdPrintBarCode.Enabled = blnWithSettingsReqDet
'
'                If ENA_SCOL_REQD_EDT = 1 Then
'                    For intRowCtr = 0 To msfgRequestDet.Rows - 1
'                        msfgRequestDet.Row = intRowCtr
'                        If Len(msfgRequestDet.TextMatrix(msfgRequestDet.Row, 8)) > 0 Then
'                            Exit For
'                        End If
'                    Next
'                Else
'                    For intRowCtr = 0 To msfgRequestDet.Rows - 1
'                        If Len(Trim(msfgRequestDet.TextMatrix(intRowCtr, 1))) = 0 And _
'                            msfgRequestDet.TextMatrix(intRowCtr, 2) = "1" And _
'                            msfgRequestDet.TextMatrix(intRowCtr, 3) = "1" Then
'                            msfgRequestDet.Row = intRowCtr
'                            Exit For
'                        End If
'                    Next
'                End If
'
'            End If
'
'            For intRowCtr = 1 To .Cols - 1
'                .Col = intRowCtr
'                .CellBackColor = &H8000000D
'                .CellForeColor = vbWhite
'            Next
'             .Enabled = False
'            msfgRequestDet.Enabled = blnWithSettingsReqDet
'            cmbSectionFilter_Click
'        End With
'
'        With msfgRequestDet
'            If .CellBackColor = &H80000005 Then
'                If UCase(.TextMatrix(.Row, 5)) = "TEXT" Then
'                    txtmsfgReqDet.Move (.CellLeft + .Left), (.CellTop + .Top), .CellWidth, .CellHeight
'                    txtmsfgReqDet.Visible = True
'                    txtmsfgReqDet.Enabled = True
'                    txtmsfgReqDet.MaxLength = setLenTxtBox(msfgRequestDet.Name, UCase(msfgRequestDet.TextMatrix(msfgRequestDet.Row, 6)))
'                    txtmsfgReqDet.TabIndex = 0
'                    txtmsfgReqDet.TabStop = True
'                    txtmsfgReqDet.Text = .TextMatrix(.Row, 1)
'                ElseIf UCase(.TextMatrix(.Row, 5)) = "COMBO" Then
'                    cmbmsfgReqDet.Visible = True
'                    cmbmsfgReqDet.Enabled = True
'                    cmbmsfgReqDet.TabIndex = 0
'                    cmbmsfgReqDet.TabStop = True
'                    cmbmsfgReqDet.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
'                    setListcmbmsfgReqDet .TextMatrix(.Row, 6)
'                    If Len(.TextMatrix(.Row, .Col)) > 0 Then
'                        For intRowCtr = 0 To cmbmsfgReqDet.ListCount - 1
'                            If .TextMatrix(.Row, 1) = cmbmsfgReqDet.List(intRowCtr) Then
'                                cmbmsfgReqDet.ListIndex = intRowCtr
'                            End If
'                        Next
'                    End If
'                End If
'            End If
'        End With
'    End If
'
'    If Me.Width > Screen.Width Then
'        sizeFix
'    End If
'
'End Sub
'
'Private Sub setListcmbmsfgReqDet(strFieldName As String)
'
'    Dim strSelectionCriteria
'
'    strSelectionCriteria = getGlobalSettings("SEL_" & strFieldName, "RE_REQ", "")
'
'    strSQLStatement = "EXEC sp_tbSelectionValues_select '" & strSelectionCriteria & "'"
'    Set rsrequest = myConnection.Execute(strSQLStatement)
'
'    With cmbmsfgReqDet
'        If .ListCount > 0 Then
'            For intRowCtr = .ListCount To 1 Step -1
'                .RemoveItem (0)
'            Next
'        End If
'        If rsrequest.RecordCount > 0 Then
'            For intRowCtr = 0 To rsrequest.RecordCount - 1
'                .AddItem (rsrequest.Fields(1))
'                rsrequest.MoveNext
'            Next
'        End If
'    End With
'
'End Sub
'
'Private Sub setInitialmsfgResult()
'    Dim strStartRowNewDemo As String, strStartRowEdtDemo As String
'
'    strStartRowNewDemo = getGlobalSettings("CUR_DEMO_NEW", "RE_DEMO", "")
'    strStartRowEdtDemo = getGlobalSettings("CUR_DEMO_EDT", "RE_DEMO", "")
'
'    With msfgResult
'        'set design and label of grid
'        If blnNewPatient Then
'            strSQLStatement = "EXEC sp_PatientDemo_select 'x'"
'        Else
'            strSQLStatement = "EXEC sp_PatientDemo_select '" & PatientID & "'"
'        End If
'
'        Set rsrequest = myConnection.Execute(strSQLStatement)
'
'        .Rows = rsrequest.Fields.Count
'
''0 = "PATIENTID"
''1 = "HOPITALID"
''2 = "LNAME"
''3 = "NAME"
''4 = "MNAME"
''5 = "GENDER"
''6 = "DOB"
''7 = "CIVILSTATUS"
''8 = "PHILID"
''9 = "ADDRESS1"
''10 = "ADDRESS2"
''11 = "EMAIL"
''12 = "PHONE"
'
'        For intRowCtr = 0 To .Rows - 1
'            .Col = 0
'            .Row = intRowCtr
'
'            .TextMatrix(.Row, 2) = rsrequest.Fields(.Row).Name
'            .TextMatrix(.Row, 0) = getGlobalSettings(rsrequest.Fields(.Row).Name, "RE_DEMO", rsrequest.Fields(.Row).Name)
'            .TextMatrix(.Row, 3) = getGlobalSettings("ENA_" & .TextMatrix(.Row, 2), "RE_DEMO", "0")
'            .TextMatrix(.Row, 4) = getGlobalSettings("VIS_" & .TextMatrix(.Row, 2), "RE_DEMO", "1")
'            .TextMatrix(.Row, 5) = getGlobalSettings("MAN_" & .TextMatrix(.Row, 2), "RE_DEMO", "1")
'            .TextMatrix(.Row, 6) = getGlobalSettings("TYPE_" & .TextMatrix(.Row, 2), "RE_DEMO", "TEXT")
'            .TextMatrix(.Row, 7) = getGlobalSettings("SEL_" & .TextMatrix(.Row, 2), "RE_DEMO", "")
'
'            If blnWithSettingsDemo = False And .TextMatrix(.Row, 3) = "1" Then
'                blnWithSettingsDemo = True
'            End If
'
'            If ENA_SCOL_DEMO_NEW = "1" Then
'                If .TextMatrix(.Row, 0) = CUR_DEMO_NEW Then
'                    .TextMatrix(.Row, 8) = "1"
'                End If
'            End If
'
'            If ENA_SCOL_DEMO_EDT = "1" Then
'                If .TextMatrix(.Row, 0) = CUR_DEMO_EDT Then
'                    .TextMatrix(.Row, 9) = "1"
'                End If
'            End If
'
'            .Col = 1
'            .CellAlignment = 1
'
'            If Not blnNewPatient Then
'                .CellBackColor = &H80000011
'                .TextMatrix(.Row, 1) = IIf(IsNull(rsrequest.Fields(.Row).Value), "", Trim(rsrequest.Fields(.Row).Value))
'                .TextMatrix(.Row, 11) = IIf(IsNull(rsrequest.Fields(.Row).Value), "", Trim(rsrequest.Fields(.Row).Value))
'            Else
'                .CellBackColor = IIf(.TextMatrix(.Row, 3) = "1", &H80000005, &H80000011)
'            End If
'            .RowHeight(.Row) = IIf(.TextMatrix(.Row, 4) = "1", 350, 0)
'        Next
'
'        For intColCtr = 0 To .Cols - 1
'            .ColWidth(intColCtr) = 0
'        Next
'
'        .ColWidth(0) = 1500
'        .ColWidth(1) = 2750
'
'    End With
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    PatientID = ""
'    RequestID = ""
'    blnNewPatient = False
'End Sub
'
'Private Sub msfgExam_Click()
'    With msfgExam
'        If .Col = 1 And (Len(.TextMatrix(.Row, 7)) = 0) And _
'            (Len(.TextMatrix(.Row, 0)) = 0 Or .TextMatrix(.Row, 0) = "0") And _
'            Len(.TextMatrix(.Row, 10)) = 0 Then
'            .CellAlignment = 0
'            txtmsfgExam.Move .CellLeft + .Left, .CellTop + .Top, .CellWidth, .CellHeight
'            txtmsfgExam.Visible = True
'            txtmsfgExam.Enabled = True
'            txtmsfgExam.Text = .TextMatrix(.Row, .Col)
'            txtmsfgExam.SetFocus
'        End If
'    End With
'End Sub
'
'Private Sub msfgExam_EnterCell()
'    If Len(msfgExam.TextMatrix(msfgExam.Row, 5)) > 0 Then
'        If msfgExam.TextMatrix(msfgExam.Row, 5) < 2 Then
'            cmdDeleteExam.Enabled = True
'        Else
'            cmdDeleteExam.Enabled = False
'        End If
'    Else
'        cmdDeleteExam.Enabled = False
'    End If
'
'    If Len(Trim(msfgExam.TextMatrix(msfgExam.Row, 5))) > 0 And _
'        Len(Trim(msfgExam.TextMatrix(msfgExam.Row, 6))) > 0 Then
'        cmdPrintBarCode.Enabled = True
'    Else
'        cmdPrintBarCode.Enabled = False
'    End If
'
'    txtmsfgExam.Visible = False
'    txtmsfgExam.Enabled = False
'
'    With msfgExam
'        If .Col = 1 And (Len(.TextMatrix(.Row, 6)) = 0) And _
'            (Len(.TextMatrix(.Row, 0)) = 0 Or .TextMatrix(.Row, 0) = "0") And _
'            Len(.TextMatrix(.Row, 9)) = 0 Then
'            .Col = 1
'            txtmsfgExam.Move .CellLeft + .Left, .CellTop + .Top, .CellWidth, .CellHeight
'            txtmsfgExam.Visible = True
'            txtmsfgExam.Enabled = True
'            txtmsfgExam.Text = .TextMatrix(.Row, .Col)
'            txtmsfgExam.SetFocus
'        End If
'    End With
'End Sub
'
'Private Sub msfgExam_GotFocus()
'    txtmsfgReqDet.Visible = False
'    cmbmsfgReqDet.Visible = False
'End Sub
'
'Private Sub msfgExam_LeaveCell()
'    With msfgExam
'        If (Len(.TextMatrix(.Row, 7)) > 0) Then
'            .TextMatrix(.Row, 1) = txtmsfgExam.Text
'        End If
'    End With
'End Sub
'
'Private Sub msfgExam_LostFocus()
'    msfgExam_LeaveCell
'End Sub
'
'Private Sub msfgExam_Scroll()
''    With msfgExam
''        If .Row = .Rows - 1 Then
''            If .RowIsVisible(.Rows - 1) And .MouseRow = .Rows - 1 And Len(.TextMatrix(.Row, 0)) > 0 Then
''                txtmsfgExam.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
''                txtmsfgExam.Visible = True
''            Else
''                txtmsfgExam.Visible = False
''            End If
''        Else
''            If .RowIsVisible(.Row) And Len(.TextMatrix(.Row, 0)) > 0 Then
''                If .Row = .Row + 1 And .RowIsVisible(.Row + 1) Then
''                    txtmsfgExam.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
''                    txtmsfgExam.Visible = True
''                Else
''                    txtmsfgExam.Visible = False
''                End If
''            Else
''                txtmsfgExam.Visible = False
''            End If
''        End If
''    End With
'End Sub
'
'Private Sub msfgRequest_Click()
'    txtmsfgResult.Visible = False
'    txtmsfgResult.Enabled = False
'    If Len(msfgRequest.TextMatrix(msfgRequest.Row, 0)) > 0 Then
'        If msfgRequest.TextMatrix(msfgRequest.Row, 0) < 2 Then
'            cmdDelReq.Enabled = True
'        Else
'            cmdDelReq.Enabled = False
'        End If
'    Else
'        cmdDelReq.Enabled = False
'    End If
'End Sub
'
'Private Sub msfgRequest_DblClick()
'    If cmdSaveReq.Caption = "Edit" And cmdSaveReq.Enabled Then
'        cmdSaveReq_Click
'    End If
'End Sub
'
'Private Sub msfgRequest_EnterCell()
'    'To Highlight Selected Request in msfgrequest'''''''''''''''''''''
'    msfgRequest.ColSel = msfgRequest.Cols - 1
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'    'Set Data of msfgrequestdet'''''''''''''''''''''''''''''''''''''''
'    If Len(msfgRequest.TextMatrix(msfgRequest.Row, 12)) > 0 Then
'
'        'Set Data'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        For intRowCtr = 0 To msfgRequestDet.Rows - 1
'            msfgRequestDet.TextMatrix(intRowCtr, 1) = ""
'            msfgRequestDet.TextMatrix(intRowCtr, 10) = ""
'            msfgRequestDet.TextMatrix(intRowCtr, 11) = ""
'            'Get Original Request By And Name of Physician''''''''''''''''''''
'            If intRowCtr = 7 Or intRowCtr = 12 Then
'                If InStrRev(msfgRequest.TextMatrix(msfgRequest.Row, intRowCtr + 1), "|") > 0 Then
'                    msfgRequestDet.TextMatrix(intRowCtr, 1) = _
'                            Mid(msfgRequest.TextMatrix(msfgRequest.Row, intRowCtr + 1), 1, _
'                            InStrRev(msfgRequest.TextMatrix(msfgRequest.Row, intRowCtr + 1), "|") - 1)
'                    msfgRequestDet.TextMatrix(intRowCtr, 10) = _
'                            Mid(msfgRequest.TextMatrix(msfgRequest.Row, intRowCtr + 1), _
'                            InStrRev(msfgRequest.TextMatrix(msfgRequest.Row, intRowCtr + 1), "|") + 1)
'                    msfgRequestDet.TextMatrix(intRowCtr, 11) = _
'                            Mid(msfgRequest.TextMatrix(msfgRequest.Row, intRowCtr + 1), 1, _
'                            InStrRev(msfgRequest.TextMatrix(msfgRequest.Row, intRowCtr + 1), "|") - 1)
'                End If
'            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'            Else
'                msfgRequestDet.TextMatrix(intRowCtr, 1) = msfgRequest.TextMatrix(msfgRequest.Row, intRowCtr + 1)
'                msfgRequestDet.TextMatrix(intRowCtr, 11) = msfgRequest.TextMatrix(msfgRequest.Row, intRowCtr + 1)
'            End If
'        Next
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'        'set the text alignment to left'''''''''''''''''''''''''''''''''''
'        For intRowCtr = 0 To msfgRequestDet.Rows - 1
'            msfgRequestDet.Row = intRowCtr
'            msfgRequestDet.CellAlignment = 0
'        Next
'        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        msfgExamDet.ClearStructure
'        msfgExamDet.Rows = 2
'
'        With msfgExam
'            Set .DataSource = Nothing
'            .Rows = 2
'            For intColCtr = 0 To .Cols - 1
'                .TextMatrix(1, intColCtr) = ""
'            Next
'            .Col = 0
'            Set .CellPicture = imglstIcon.ListImages(6).Picture
'            .CellAlignment = 4
'            .Enabled = False
'        End With
'
'        setmsfgExam Trim(msfgRequest.TextMatrix(msfgRequest.Row, 6))
'        msfgExam.Enabled = False
'
'        If Val(msfgRequest.TextMatrix(msfgRequest.Row, 12)) > 2 Then
'            cmdSaveReq.Enabled = False
'
'            With msfgExam
'                If .Rows > 2 Then
'                    .RemoveItem (.Rows - 1)
'                End If
'            End With
'
'        Else
'            cmdSaveReq.Enabled = blnWithSettingsReqDet
'        End If
'
'        cmdSaveReq.Caption = "Edit"
'    End If
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'    'Set Endable of cmdCopyRequest''''''''''''''''''''''''''''''''''''
'    If Len(msfgRequest.TextMatrix(msfgRequest.Row, 12)) > 0 Then
'        cmdCopyRequest.Enabled = blnWithSettingsReqDet
'    Else
'        cmdCopyRequest.Enabled = False
'    End If
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
''    'Set Endable of cmddelreq button''''''''''''''''''''''''''''''''''
''    If Len(msfgRequest.TextMatrix(msfgRequest.Row, 12)) > "0" Then
''        If msfgRequest.TextMatrix(msfgRequest.Row, 12) < 2 Then
''            cmdDelReq.Enabled = blnWithSettingsReqDet
''        Else
''            cmdDelReq.Enabled = False
''        End If
''    Else
''        cmdDelReq.Enabled = False
''    End If
''    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'End Sub
'
'Private Sub msfgRequest_KeyDown(KeyCode As Integer, Shift As Integer)
'    With msfgRequest
'        Select Case KeyCode
'            Case vbKeyUp
'                If .Row > 1 Then
'                    If .RowIsVisible(.Row - 1) = False Then
'                        .TopRow = .Row
'                    End If
'                Else
'                    .TopRow = .Row
'                End If
'            Case vbKeyDown
'                If .Row < .Rows - 1 Then
'                    If .RowIsVisible(.Row + 1) = False Then
'                        .TopRow = .Row
'                    End If
'                Else
'                    .TopRow = .Row
'                End If
'            Case vbKeyPageUp
'                .TopRow = .Row
'            Case vbKeyPageDown
'                If .Row < .Rows - 1 Then
'                    If .RowIsVisible(.Row + 1) = False Then
'                        .TopRow = .Row
'                    End If
'                Else
'                    .TopRow = .Row
'                End If
'        End Select
'    End With
'End Sub
'
'Private Sub msfgRequest_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If cmdSaveReq.Caption = "Edit" And cmdSaveReq.Enabled Then
'            cmdSaveReq_Click
'        End If
'    End If
'End Sub
'
'Private Sub msfgRequest_LeaveCell()
'    For intRowCtr = 1 To msfgRequest.Cols - 1
'        msfgRequest.Col = intRowCtr
'        msfgRequest.CellBackColor = &H80000005
'        msfgRequest.CellForeColor = vbBlack
'    Next
'    msfgRequest.Col = 0
'End Sub
'
'Private Sub msfgRequestDet_Click()
'    With msfgRequestDet
'        .CellAlignment = 0
'        If Len(.TextMatrix(12, 1)) > 0 And Len(.TextMatrix(12, 10)) = 0 And .TextMatrix(12, 5) <> "COMBO" Then
'            .Row = 12
'            txtmsfgReqDet.SetFocus
'            highLight
'        Else
'            If .TextMatrix(.Row, 2) = "1" And .TextMatrix(.Row, 3) = "1" And msfgRequest.TextMatrix(msfgRequest.Row, 0) < "2" Then
'                If UCase(.TextMatrix(.Row, 5)) = "TEXT" Then
'                    txtmsfgReqDet.Move (.CellLeft + .Left), (.CellTop + .Top), .CellWidth, .CellHeight
'                    txtmsfgReqDet.Visible = True
'                    txtmsfgReqDet.Enabled = True
'                    txtmsfgReqDet.MaxLength = setLenTxtBox(msfgRequestDet.Name, UCase(.TextMatrix(.Row, 6)))
'                    txtmsfgReqDet.Text = .TextMatrix(.Row, 1)
'                    cmbmsfgReqDet.Visible = False
'                    txtmsfgReqDet.SetFocus
'                ElseIf UCase(.TextMatrix(.Row, 5)) = "COMBO" Then
'                    cmbmsfgReqDet.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
'                    cmbmsfgReqDet.Visible = True
'                    cmbmsfgReqDet.Enabled = True
'                    setListcmbmsfgReqDet .TextMatrix(.Row, 6)
'                    If Len(.TextMatrix(.Row, .Col)) > 0 Then
'                        For intRowCtr = 0 To cmbmsfgReqDet.ListCount - 1
'                            If .TextMatrix(.Row, 1) = cmbmsfgReqDet.List(intRowCtr) Then
'                                cmbmsfgReqDet.ListIndex = intRowCtr
'                            End If
'                        Next
'                    End If
'                    txtmsfgReqDet.Visible = False
'                    cmbmsfgReqDet.SetFocus
'                End If
'            End If
'        End If
'    End With
'End Sub
'
'Private Sub msfgRequestDet_GotFocus()
'    msfgRequestDet_Click
'
'    If txtmsfgExam.Visible And txtmsfgExam.Enabled Then
'        txtmsfgExam.Visible = False
'        txtmsfgExam.Enabled = False
'    End If
'End Sub
'
'Private Sub msfgRequestDet_LeaveCell()
'    With msfgRequestDet
'        If .TextMatrix(.Row, 9) = "RequestPhysicianID" And UCase(.TextMatrix(.Row, 5)) <> "COMBO" And _
'            .TextMatrix(.Row, 2) = "1" And .TextMatrix(.Row, 3) = "1" Then
'             If Len(Trim(txtmsfgReqDet.Text)) > 0 Then
'
'                If IsNumeric(txtmsfgReqDet.Text) Then
'                    strSQLStatement = "EXEC sp_SearchCodes '1','" & txtmsfgReqDet.Text & "',''"
'                Else
'                    strSQLStatement = "EXEC sp_SearchCodes '1','','" & Replace(txtmsfgReqDet.Text, "'", "''") & "'"
'                End If
'
'                Set rsrequest = myConnection.Execute(strSQLStatement)
'
'                If rsrequest.RecordCount > 0 Then
'                    txtmsfgReqDet.Text = rsrequest.Fields(1)
'                    .TextMatrix(.Row, 1) = txtmsfgReqDet.Text
'                    msfgRequestDet.TextMatrix(msfgRequestDet.Row, 10) = rsrequest.Fields(0)
'
'                    If msfgRequestDet.Row < msfgRequestDet.Rows - 1 Then
'                        msfgRequestDet.Col = 0
'                        msfgRequestDet.CellFontBold = False
'                        msfgRequestDet.Col = 1
'                        msfgRequestDet_Click
'                    Else
'                        txtmsfgReqDet.SetFocus
'                    End If
'                Else
'                    .TextMatrix(.Row, 1) = txtmsfgReqDet.Text
'                    .Row = 12
'                    highLight
'                End If
'            ElseIf Len(Trim(txtmsfgReqDet.Text)) = 0 And Len(msfgRequestDet.TextMatrix(msfgRequestDet.Row, 10)) = 0 Then
'                If msfgRequestDet.Row < msfgRequestDet.Rows - 1 Then
'                    msfgRequestDet.Col = 0
'                    msfgRequestDet.CellFontBold = False
'                    msfgRequestDet.Col = 1
'                    .TextMatrix(.Row, 1) = ""
'                    msfgRequestDet_Click
'                Else
'                    txtmsfgReqDet.SetFocus
'                End If
'            End If
'        Else
'            If .CellBackColor <> &H80000011 Then
'                If UCase(.TextMatrix(.Row, 5)) = "COMBO" Then
'                    .TextMatrix(.Row, .Col) = cmbmsfgReqDet.Text
'                ElseIf UCase(.TextMatrix(.Row, 5)) = "TEXT" Then
'                    .TextMatrix(.Row, .Col) = txtmsfgReqDet.Text
'                End If
'            End If
'        End If
'    End With
'End Sub
'
'Private Sub msfgRequestDet_Scroll()
'    Dim scrollerr
'    With msfgRequestDet
'        On Error GoTo scrollerr
'        If Len(msfgRequest.Text) = 0 Or msfgRequest.TextMatrix(msfgRequest.Row, 0) = "0" Or _
'            msfgRequest.TextMatrix(msfgRequest.Row, 0) = "1" Then
'            If .Row = .Rows - 1 Then
'                If Len(txtmsfgReqDet.Text) > 0 And Len(msfgRequestDet.TextMatrix(12, 10)) = 0 Then
'                    .Row = 12
'                End If
'                If .RowIsVisible(.Rows - 1) And .MouseRow = .Rows - 1 Then
'                    If UCase(.TextMatrix(.Row, 5)) = "TEXT" Then
'                        txtmsfgReqDet.Move (.CellLeft + .Left), (.CellTop + .Top), .CellWidth, .CellHeight
'                        txtmsfgReqDet.Visible = True
'                        txtmsfgReqDet.SetFocus
'                    ElseIf UCase(.TextMatrix(.Row, 5)) = "COMBO" Then
'                        cmbmsfgReqDet.Visible = True
'                        cmbmsfgReqDet.Move .Left + .CellLeft, .Top + .CellTop
'                        cmbmsfgReqDet.SetFocus
'                    End If
'                Else
'                    If UCase(.TextMatrix(.Row, 5)) = "TEXT" Then
'                        txtmsfgReqDet.Visible = False
'                    ElseIf UCase(.TextMatrix(.Row, 5)) = "COMBO" Then
'                        cmbmsfgReqDet.Visible = False
'                    End If
'                End If
'            Else
'                If .Row > 12 Then
'                    If Len(msfgRequestDet.TextMatrix(12, 1)) > 0 And Len(msfgRequestDet.TextMatrix(12, 10)) = 0 Then
'                        .Row = 12
'                    End If
'                End If
'                If .RowIsVisible(.Row + 1) And .RowIsVisible(.Row - 1) Then
'                   If UCase(.TextMatrix(.Row, 5)) = "TEXT" Then
'                       txtmsfgReqDet.Move (.CellLeft + .Left), (.CellTop + .Top), .CellWidth, .CellHeight
'                       txtmsfgReqDet.Visible = True
'                       txtmsfgReqDet.SetFocus
'                   ElseIf UCase(.TextMatrix(.Row, 5)) = "COMBO" Then
'                       cmbmsfgReqDet.Visible = True
'                       cmbmsfgReqDet.Move .Left + .CellLeft, .Top + .CellTop
'                       cmbmsfgReqDet.SetFocus
'                   End If
'                Else
'                    If UCase(.TextMatrix(.Row, 5)) = "TEXT" Then
'                       txtmsfgReqDet.Visible = False
'                   ElseIf UCase(.TextMatrix(.Row, 5)) = "COMBO" Then
'                       cmbmsfgReqDet.Visible = False
'                   End If
'                End If
'            End If
'        End If
'    End With
'scrollerr:
'    If err.Number = 30009 Then
'        msfgRequestDet.TopRow = 0
'    Else
'        Resume Next
'    End If
'
'End Sub
'
'Private Sub msfgResult_Click()
'    With msfgResult
'        If .Enabled = True Then
'            If .Col = 1 Then
'                If .TextMatrix(.Row, 3) = "1" And .TextMatrix(.Row, 4) = "1" Then
'                    .CellAlignment = 0
'                    If .TextMatrix(.Row, 6) = "TEXT" Then
'                        txtmsfgResult.Move (.CellLeft + .Left), (.CellTop + .Top), .CellWidth, .CellHeight
'                        txtmsfgResult.Visible = True
'                        txtmsfgResult.Enabled = True
'                        txtmsfgResult.Text = .TextMatrix(.Row, .Col)
'                        txtmsfgResult.MaxLength = setLenTxtBox(msfgResult.Name, .TextMatrix(.Row, 3))
'                        txtmsfgResult.SetFocus
'                    Else
'                        cmbmsfgResult.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth
'                        cmbmsfgResult.Visible = True
'                        cmbmsfgResult.Enabled = True
'                        setListcmbmsfgReqDet .TextMatrix(.Row, 6)
'                        If Len(.TextMatrix(.Row, .Col)) > 0 Then
'                            For intRowCtr = 0 To cmbmsfgResult.ListCount - 1
'                                If .TextMatrix(.Row, 1) = cmbmsfgResult.List(intRowCtr) Then
'                                    cmbmsfgResult.ListIndex = intRowCtr
'                                End If
'                            Next
'                        End If
'                        txtmsfgResult.Visible = False
'                        cmbmsfgResult.SetFocus
'                    End If
'                Else
'                    If txtmsfgResult.Visible = True And txtmsfgResult.Enabled = True Then
'                        txtmsfgResult.SetFocus
'                    ElseIf cmbmsfgResult.Visible = True And cmbmsfgResult.Enabled = True Then
'                        cmbmsfgResult.SetFocus
'                    End If
'                End If
'            Else
'                If txtmsfgResult.Visible = True And txtmsfgResult.Enabled = True Then
'                    txtmsfgResult.SetFocus
'                ElseIf cmbmsfgResult.Visible = True And cmbmsfgResult.Enabled = True Then
'                    cmbmsfgResult.SetFocus
'                End If
'            End If
'        End If
'    End With
'End Sub
'
'Private Sub msfgResult_GotFocus()
'    msfgResult_Click
'End Sub
'
'Private Sub msfgResult_LeaveCell()
'    With msfgResult
'        If .TextMatrix(.Row, 3) = "1" And .TextMatrix(.Row, 4) = "1" Then
'            If .TextMatrix(.Row, 6) = "TEXT" Then
'                .TextMatrix(.Row, .Col) = txtmsfgResult.Text
'            Else
'                .TextMatrix(.Row, .Col) = cmbmsfgResult.Text
'            End If
'        End If
'    End With
'End Sub
'
'Private Sub msfgResult_Scroll()
'    txtmsfgResult.Visible = False
'    txtmsfgResult.Enabled = False
'    cmbmsfgResult.Visible = False
'    cmbmsfgResult.Enabled = False
'End Sub
'
'Private Sub txtmsfgExam_Change()
'    txtmsfgExam.Text = UCase(txtmsfgExam.Text)
'    txtmsfgExam.SelStart = Len(txtmsfgExam.Text)
'End Sub
'
'Private Sub txtmsfgExam_GotFocus()
'    highLight
'    txtmsfgExam.BackColor = &H80FF80
'End Sub
'
'Private Sub txtmsfgExam_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim strSelectionValue As String
'    If KeyCode = 113 Then
'
'        strSelectionValue = showfrmExamSelection(txtmsfgExam.Text, getValue(1, getUserDet(strRequestBy)), cmbSectionFilter.Text)
'
'        If Len(strSelectionValue) > 0 Then
''0=status
''1=examcode
''2=instrumentid
''3=rvid
''4=requestid
''5=sectionid
''6=status
''7=sampleid
''8=specimen
''9=examname
'            If Not blnExamOccured(strSelectionValue) Then
'                With msfgExam
'                    .TextMatrix(.Row, 0) = "0"
'                    .CellAlignment = 1
'                    .TextMatrix(.Row, 1) = getValue(1, strSelectionValue)
'                    .TextMatrix(.Row, 2) = getValue(6, strSelectionValue)
'                    .TextMatrix(.Row, 3) = msfgRequest.TextMatrix(msfgRequest.Row, 6)
'                    .TextMatrix(.Row, 4) = getValue(5, strSelectionValue)
'                    .TextMatrix(.Row, 5) = "0"
'                    .TextMatrix(.Row, 6) = ""
'                    .TextMatrix(.Row, 7) = getValue(3, strSelectionValue)
'                    .TextMatrix(.Row, 8) = getValue(2, strSelectionValue)
'                    .TextMatrix(.Row, 9) = ""
'                    .TextMatrix(.Row, 10) = getValue(4, strSelectionValue)
'                    .Col = 0
'                    .CellForeColor = vbWhite
'
'                    AddExamDet .Row
'
'                    If Len(.TextMatrix(.Rows - 1, 0)) > 0 Then
'                        .Rows = .Rows + 1
'                    End If
'                    .Row = .Rows - 1
'
'                    .Col = 0
'                    .CellForeColor = vbWhite
'                    .CellBackColor = vbWhite
'
'                    Set .CellPicture = imglstIcon.ListImages(6).Picture
'                    .CellPictureAlignment = 4
'                    .Col = 1
'
'                    .RowHeight(.Row) = 350
'                    txtmsfgExam.Move .Left + .CellLeft, .Top + .CellTop
'                    txtmsfgExam.Visible = True
'                    txtmsfgExam.Text = ""
'                    .CellAlignment = 0
'                End With
'            Else
'                MsgBox "Unable to Add " & getValue(2, strSelectionValue)
'            End If
'        End If
'        If txtmsfgExam.Enabled Then
'            txtmsfgExam.SetFocus
'        End If
'    ElseIf KeyCode = 38 Then
'        With msfgExam
'            .TopRow = .Row
'            If .Row > 1 Then
'                msfgExam_LeaveCell
'                .Row = .Row - 1
'                .SetFocus
'                .ColSel = .Cols - 1
'                msfgExam_EnterCell
'            End If
'        End With
'    ElseIf KeyCode = 40 Then
'        With msfgExam
'             .TopRow = .Row
'            If .Row < .Rows - 1 Then
'                msfgExam_LeaveCell
'                .Row = .Row + 1
'                .SetFocus
'                msfgExam_EnterCell
'                highLight
'            End If
'        End With
'    ElseIf KeyCode = vbKeyEscape Then
'        txtmsfgExam.Visible = False
'        cmdCancelReq_Click
'    Else
'        If Len(msfgExam.TextMatrix(msfgExam.Row, 1)) > 0 Then
'            DelExamDet msfgExam.TextMatrix(msfgExam.Row, 1)
'        End If
'        For intRowCtr = 0 To msfgExam.Cols - 1
'            msfgExam.TextMatrix(msfgExam.Row, intRowCtr) = ""
'        Next
'    End If
'End Sub
'
'Private Sub txtmsfgExam_KeyPress(KeyAscii As Integer)
'    Dim strSelectionValue As String
'
'    If KeyAscii = 13 Then
'
'        If Len(Trim(txtmsfgExam.Text)) = 0 Then
'            strSelectionValue = showfrmExamSelection(txtmsfgExam.Text, getValue(1, getUserDet(strRequestBy)), cmbSectionFilter.Text)
'        Else
'            strSQLStatement = "EXEC sp_ReqExam_Select NULL," & _
'                                "'" & Replace(txtmsfgExam.Text, "'", "''") & "'"
''            If getValue(1, getUserDet(strRequestBy)) = 4 Then
''                strSQLStatement = "EXEC sp_ReqExam_Select NULL," & _
''                                                "NULL,'" & Replace(txtmsfgExam.Text, "'", "''") & "'"
''            Else
''                strSQLStatement = "EXEC sp_ReqExam_Select " & IIf(Len(InstrumentID) > 0, "'" & InstrumentID & "'", "NULL") & "," & _
''                                                "NULL,'" & Replace(txtmsfgExam.Text, "'", "''") & "'"
''            End If
'
'            Set rsrequest = myConnection.Execute(strSQLStatement)
'            If rsrequest.RecordCount > 0 Then
'                If rsrequest.RecordCount > 1 Then
'                    strSelectionValue = showfrmExamSelection(txtmsfgExam.Text, getValue(1, getUserDet(strRequestBy)), cmbSectionFilter.Text)
'                Else
'                    For intColCtr = 0 To rsrequest.Fields.Count - 1
'                        If Len(strSelectionValue) = 0 Then
'                            strSelectionValue = "|" & rsrequest.Fields(intColCtr) & "|"
'                        Else
'                            strSelectionValue = strSelectionValue & rsrequest.Fields(intColCtr) & "|"
'                        End If
'                    Next
'                End If
'            Else
'                strSelectionValue = showfrmExamSelection(txtmsfgExam.Text, getValue(1, getUserDet(strRequestBy)), cmbSectionFilter.Text)
'            End If
'        End If
'        If Len(strSelectionValue) > 0 Then
'            If Not blnExamOccured(strSelectionValue) Then
'                With msfgExam
'                    .TextMatrix(.Row, 0) = "0"
'                    .CellAlignment = 1
'                    .TextMatrix(.Row, 1) = getValue(1, strSelectionValue)
'                    .TextMatrix(.Row, 2) = getValue(6, strSelectionValue)
'                    .TextMatrix(.Row, 3) = msfgRequest.TextMatrix(msfgRequest.Row, 6)
'                    .TextMatrix(.Row, 4) = getValue(5, strSelectionValue)
'                    .TextMatrix(.Row, 5) = "0"
'                    .TextMatrix(.Row, 6) = ""
'                    .TextMatrix(.Row, 7) = getValue(3, strSelectionValue)
'                    .TextMatrix(.Row, 8) = getValue(2, strSelectionValue)
'                    .TextMatrix(.Row, 9) = ""
'                    .TextMatrix(.Row, 10) = getValue(4, strSelectionValue)
'                    .Col = 0
'                    .CellForeColor = vbWhite
'
'                    AddExamDet .Row
'
'                    If Len(.TextMatrix(.Rows - 1, 0)) > 0 Then
'                        .Rows = .Rows + 1
'                    End If
'
'                    .Row = .Rows - 1
'
'                    .Col = 0
'                    .CellBackColor = vbWhite
'                    .CellForeColor = vbBlack
'
'                    Set .CellPicture = imglstIcon.ListImages(6).Picture
'                    .CellPictureAlignment = 4
'                    .Col = 1
'
'                    .RowHeight(.Row) = 350
'                    txtmsfgExam.Move .Left + .CellLeft, .Top + .CellTop
'                    txtmsfgExam.Visible = True
'                    txtmsfgExam.Text = ""
'                    .CellAlignment = 0
'
'                End With
'            Else
'                MsgBox "Unable to Add " & getValue(2, strSelectionValue)
'            End If
'            KeyAscii = 0
'        End If
'        If txtmsfgExam.Enabled Or txtmsfgExam.Visible Then
'            txtmsfgExam.SetFocus
'        End If
'    End If
'
'    cmdDeleteExam.Enabled = False
'End Sub
'
'Private Sub txtmsfgReqDet_Change()
'    txtmsfgReqDet.Text = UCase(txtmsfgReqDet.Text)
'    txtmsfgReqDet.SelStart = Len(txtmsfgReqDet.Text)
'
'    If msfgRequestDet.TextMatrix(msfgRequestDet.Row, 9) = "RequestPhysicianID" Then
'        If msfgRequestDet.TextMatrix(msfgRequestDet.Row, 1) <> txtmsfgReqDet.Text Then
'            msfgRequestDet.TextMatrix(msfgRequestDet.Row, 10) = ""
'            msfgRequestDet.Col = 0
'            msfgRequestDet.CellFontBold = True
'            msfgRequestDet.Col = 1
'        End If
'    End If
'End Sub
'
'Private Sub txtmsfgReqDet_GotFocus()
'    highLight
'    highLight_bg
'End Sub
'
'Private Sub txtmsfgReqDet_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 38 Then
'        If msfgRequestDet.Row > 1 Then
'            msfgRequestDet_LeaveCell
'            If msfgRequestDet.TextMatrix(msfgRequestDet.Row, 9) = "RequestPhysicianID" Then
'                With msfgRequestDet
'                    If Len(txtmsfgReqDet.Text) > 0 And Len(.TextMatrix(.Row, 10)) > 0 Then
'                        For intRowCtr = msfgRequestDet.Row - 1 To 0 Step -1
'                            If msfgRequestDet.TextMatrix(intRowCtr, 2) = "1" And _
'                                msfgRequestDet.TextMatrix(intRowCtr, 3) = "1" Then
'                                msfgRequestDet.Row = intRowCtr
'                                Exit For
'                            End If
'                        Next
'                    ElseIf Len(txtmsfgReqDet.Text) = 0 And Len(.TextMatrix(.Row, 10)) = 0 Then
'                        For intRowCtr = msfgRequestDet.Row - 1 To 0 Step -1
'                                If msfgRequestDet.TextMatrix(intRowCtr, 2) = "1" And _
'                                    msfgRequestDet.TextMatrix(intRowCtr, 3) = "1" Then
'                                    msfgRequestDet.Row = intRowCtr
'                                    Exit For
'                                End If
'                            Next
'                        End If
'                End With
'            Else
'                For intRowCtr = msfgRequestDet.Row - 1 To 0 Step -1
'                    If msfgRequestDet.TextMatrix(intRowCtr, 2) = "1" And _
'                        msfgRequestDet.TextMatrix(intRowCtr, 3) = "1" Then
'                        msfgRequestDet.Row = intRowCtr
'                        Exit For
'                    End If
'                Next
'            End If
'            msfgRequestDet.SetFocus
'            msfgRequestDet_Click
'        Else
'            msfgRequestDet.TopRow = 0
'            txtmsfgReqDet.SetFocus
'        End If
'    ElseIf KeyCode = 40 Then
'        If msfgRequestDet.Row < msfgRequestDet.Rows - 1 Then
'            msfgRequestDet_LeaveCell
'            If msfgRequestDet.TextMatrix(msfgRequestDet.Row, 9) = "RequestPhysicianID" Then
'                With msfgRequestDet
'                    If Len(txtmsfgReqDet.Text) > 0 And Len(.TextMatrix(.Row, 10)) > 0 Then
'                        For intRowCtr = msfgRequestDet.Row + 1 To msfgRequestDet.Rows - 1
'                            If msfgRequestDet.TextMatrix(intRowCtr, 2) = "1" And _
'                                    msfgRequestDet.TextMatrix(intRowCtr, 3) = "1" Then
'                                msfgRequestDet.Row = intRowCtr
'                                Exit For
'                            End If
'                        Next
'                    ElseIf Len(txtmsfgReqDet.Text) = 0 And Len(.TextMatrix(.Row, 10)) = 0 Then
'                        For intRowCtr = msfgRequestDet.Row + 1 To msfgRequestDet.Rows - 1
'                            If msfgRequestDet.TextMatrix(intRowCtr, 2) = "1" And _
'                                    msfgRequestDet.TextMatrix(intRowCtr, 3) = "1" Then
'                                msfgRequestDet.Row = intRowCtr
'                                Exit For
'                            End If
'                        Next
'                    End If
'                End With
'            Else
'                For intRowCtr = msfgRequestDet.Row + 1 To msfgRequestDet.Rows - 1
'                    If msfgRequestDet.TextMatrix(intRowCtr, 2) = "1" And _
'                            msfgRequestDet.TextMatrix(intRowCtr, 3) = "1" Then
'                        msfgRequestDet.Row = intRowCtr
'                        Exit For
'                    End If
'                Next
'            End If
'            msfgRequestDet.SetFocus
'            msfgRequestDet_Click
'            highLight
'        Else
'            txtmsfgReqDet.SetFocus
'        End If
'    ElseIf KeyCode = 113 And msfgRequestDet.Row = 12 Then
'        Dim strReturn As String
'        strReturn = searchObj.showSearchCode("1", txtmsfgReqDet.Text, False, , True, True)
'        If Not isExiting Then
'            msfgRequestDet.TextMatrix(msfgRequestDet.Row, 10) = ""
'            If Len(Trim(strReturn)) > 0 Then
'                txtmsfgReqDet.SetFocus
'                msfgRequestDet.TextMatrix(msfgRequestDet.Row, 10) = Mid(strReturn, 1, InStr(1, strReturn, "|") - 1)
'                msfgRequestDet.TextMatrix(msfgRequestDet.Row, 1) = Mid(strReturn, InStr(1, strReturn, "|") + 1)
'                txtmsfgReqDet.Text = Mid(strReturn, InStr(1, strReturn, "|") + 1)
'            End If
'            txtmsfgReqDet.SetFocus
'        End If
'    ElseIf KeyCode = vbKeyEscape Then
'        cmdCancelReq_Click
'    End If
'End Sub
'
'Private Sub txtmsfgReqDet_KeyPress(KeyAscii As Integer)
'    Dim intCurrentRowTxt As Integer
'    If KeyAscii = 13 Then
'        If msfgRequestDet.Row < msfgRequestDet.Rows - 1 Then
'            msfgRequestDet_LeaveCell
'            intCurrentRowTxt = msfgRequestDet.Row
'            For intRowCtr = msfgRequestDet.Row + 1 To msfgRequestDet.Rows - 1
'                If msfgRequestDet.TextMatrix(intRowCtr, 2) = "1" And _
'                        msfgRequestDet.TextMatrix(intRowCtr, 3) = "1" Then
'                    msfgRequestDet.Row = intRowCtr
'                    Exit For
'                End If
'            Next
'            If msfgRequestDet.Row = intCurrentRowTxt Then
'                txtmsfgReqDet.Visible = False
'                txtmsfgReqDet.Enabled = False
'                If cmdSaveReq.Caption = "&Update" And msfgExam.Enabled Then
'                    msfgExam.Col = 1
'                    msfgExam_Click
'                Else
'                    cmdSaveReq.SetFocus
'                End If
'            Else
'                msfgRequestDet_Click
'                highLight
'            End If
'        Else
'            txtmsfgReqDet.SetFocus
'        End If
'        KeyAscii = 0
'    End If
'End Sub
'
'Private Sub txtmsfgResult_Change()
'    txtmsfgResult.Text = UCase(txtmsfgResult.Text)
'    txtmsfgResult.SelStart = Len(txtmsfgResult.Text)
'End Sub
'
'Private Sub txtmsfgResult_GotFocus()
'    highLight
'    highLight_bg
'End Sub
'
'Private Sub txtmsfgResult_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyUp Then
'        If msfgResult.Row > 0 Then
'            msfgResult_LeaveCell
'            For intRowCtr = msfgResult.Row - 1 To 0 Step -1
'                If msfgResult.TextMatrix(intRowCtr, 3) = "1" And msfgResult.TextMatrix(intRowCtr, 4) = "1" Then
'                    msfgResult.Row = intRowCtr
'                    Exit For
'                End If
'            Next
'
'            If msfgResult.RowIsVisible(3) Then
'                msfgResult.TopRow = 0
'            End If
'
'            msfgResult.SetFocus
'            msfgResult_Click
'        Else
'            txtmsfgResult.SetFocus
'            highLight
'        End If
'
'    ElseIf KeyCode = vbKeyDown Then
'        If msfgResult.Row < msfgResult.Rows - 1 Then
'            msfgResult_LeaveCell
'            For intRowCtr = msfgResult.Row + 1 To msfgResult.Rows - 1
'                If msfgResult.TextMatrix(intRowCtr, 3) = "1" And msfgResult.TextMatrix(intRowCtr, 4) = "1" Then
'                    msfgResult.Row = intRowCtr
'                    Exit For
'                End If
'            Next
'            msfgResult.SetFocus
'            msfgResult_Click
'            highLight
'        Else
'            txtmsfgResult.SetFocus
'            highLight
'        End If
'    End If
'End Sub
'
'Private Sub txtmsfgResult_KeyPress(KeyAscii As Integer)
'    Dim msfgResultCurrentRow As Integer
'    If KeyAscii = 13 Then
'        msfgResultCurrentRow = msfgResult.Row
'        msfgResult_LeaveCell
'        For intRowCtr = msfgResult.Row + 1 To msfgResult.Rows - 1
'            If msfgResult.TextMatrix(intRowCtr, 3) = "1" And msfgResult.TextMatrix(intRowCtr, 4) = "1" Then
'                msfgResult.Row = intRowCtr
'                Exit For
'            End If
'        Next
'        If msfgResult.Row = msfgResultCurrentRow Then
'            If cmdSaveDemo.Enabled Then
'                cmdSaveDemo.SetFocus
'                txtmsfgResult.Visible = False
'                txtmsfgResult.Enabled = False
'            End If
'        Else
'            msfgResult.SetFocus
'            msfgResult_Click
'            highLight
'        End If
'        KeyAscii = 0
'    End If
'
'    If msfgResult.TextMatrix(msfgResult.Row, 2) = "PGENDER" Then
'        KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        If InStr(1, "MF", Chr(KeyAscii)) > 0 Then
'            If Chr(KeyAscii) = "M" Then txtmsfgResult.Text = "MALE"
'            If Chr(KeyAscii) = "F" Then txtmsfgResult.Text = "FEMALE"
'        End If
'        If KeyAscii = 32 Then
'            If txtmsfgResult.Text = "MALE" Then
'                txtmsfgResult.Text = "FEMALE"
'            ElseIf txtmsfgResult.Text = "FEMALE" Then
'                txtmsfgResult.Text = "MALE"
'            Else
'                txtmsfgResult.Text = "MALE"
'            End If
'        End If
'        KeyAscii = 0
'    End If
'
'End Sub
'
'Private Function setLenTxtBox(msfgName As String, columnName As String) As Integer
'    Select Case msfgName
'        Case "msfgResult"
'            Select Case columnName
'                Case "PID"
'                    setLenTxtBox = 10
'                Case "LNAME"
'                    setLenTxtBox = 50
'                Case "NAME"
'                    setLenTxtBox = 50
'                Case "MNAME"
'                    setLenTxtBox = 50
'                Case "GENDER"
'                    setLenTxtBox = 10
'                Case "DOB"
'                    setLenTxtBox = 10
'                Case "CS"
'                    setLenTxtBox = 10
'                Case "PHLID"
'                    setLenTxtBox = 20
'                Case "ADDRESS"
'                    setLenTxtBox = 50
'                Case "ADDRESS2"
'                    setLenTxtBox = 50
'                Case "EMAIL"
'                    setLenTxtBox = 50
'                Case "PHONE"
'                    setLenTxtBox = 20
'            End Select
'        Case "msfgRequestDet"
'            Select Case columnName
'                Case "DATE"
'                    setLenTxtBox = 20
'                Case "WARD"
'                    setLenTxtBox = 25
'                Case "ROOM"
'                    setLenTxtBox = 25
'                Case "PTYPE"
'                    setLenTxtBox = 10
'                Case "LOCATION"
'                    setLenTxtBox = 20
'                Case "ID"
'                    setLenTxtBox = 10
'                Case "PID"
'                    setLenTxtBox = 20
'                Case "BY"
'                    setLenTxtBox = 10
'                Case "AGE"
'                    setLenTxtBox = 20
'                Case "PRIORITY"
'                    setLenTxtBox = 10
'                Case "COMMENT"
'                    setLenTxtBox = 150
'                Case "STATUS"
'                    setLenTxtBox = 10
'                Case "PHYSICIANID"
'                    setLenTxtBox = 60
'                Case "RESERVE1"
'                    setLenTxtBox = 20
'                Case "RESERVE2"
'                    setLenTxtBox = 20
'                Case "RESERVE3"
'                    setLenTxtBox = 20
'                Case "RESREVE4"
'                    setLenTxtBox = 20
'                Case "RESERVE5"
'                    setLenTxtBox = 20
'            End Select
'    End Select
'End Function
'
'Private Sub setInitialmsfgRequest()
'    With msfgRequest
'        Set rsSearch = myConnection.Execute("exec 2 '" & PatientID & "'")
'
'        If rsSearch.RecordCount > 0 Then
'            Set .DataSource = rsSearch
'            .Enabled = True
'        Else
'            .CellAlignment = 0
'            .Cols = rsSearch.Fields.Count + 1
'            .Enabled = False
'        End If
'
''0=status
''1=requestdate
''2=ward
''3=room
''4=patienttype
''5=location
''6=requestid
''7=requestpid
''8=requestby
''9=age
''10=priority
''11=communication
''12=status
''13=physician
''14=reserve1
''15=reserve2
''16=reserve3
''17=reserve4
''18=reserve5
''19=requetByShow
'
'        .Cols = .Cols + 1
'
'        '.TextArray(0) = "Staus"
'        .TextArray(1) = "Request Date"
'        .TextArray(2) = "Ward"
'        .TextArray(3) = "Room"
'        .TextArray(4) = "Patient Type"
'        .TextArray(19) = "Request By"
'
'        For intRowCtr = 0 To .Rows - 1
'            .RowHeight(intRowCtr) = 350
'            If InStr(1, .TextMatrix(intRowCtr, 8), "|") > 0 Then
'                .TextMatrix(intRowCtr, 19) = Mid(.TextMatrix(intRowCtr, 8), 1, _
'                                InStr(1, .TextMatrix(intRowCtr, 8), "|") - 1)
'            Else
'                .TextMatrix(intRowCtr, 19) = .TextMatrix(intRowCtr, 8)
'            End If
'        Next
'
'        For intColCtr = 0 To .Cols - 1
'            .ColWidth(intColCtr) = 0
'        Next
'
'        .ColWidth(0) = 600
'        .ColWidth(1) = 2500
'        .ColWidth(2) = 1800
'        .ColWidth(3) = 1800
'        .ColWidth(4) = 1800
'        .ColWidth(19) = 2000
'
''        For intRowCtr = 1 To .Rows - 1
''            .TextMatrix(intRowCtr, 0) = .TextMatrix(intRowCtr, 12)
''        Next
'
'    End With
'
'End Sub
'
'Private Sub setInitialmsfgRequestDet()
'    Dim strStartRowNewReq As String, strStartRowEdtReq As String
'
'    strStartRowNewReq = getGlobalSettings("CUR_REQD_NEW", "RE_REQ", "")
'    strStartRowEdtReq = getGlobalSettings("CUR_REQD_EDT", "RE_REQ", "")
'
'    With msfgRequestDet
'        strSQLStatement = "exec 2 NULL"
'        Set rsrequest = myConnection.Execute(strSQLStatement)
'
'        .Rows = rsrequest.Fields.Count
'
'        'set design and label of grid
'        For intRowCtr = 0 To .Rows - 1
'            If Left(rsrequest.Fields(intRowCtr).Name, 7) = "Request" Then
'                .TextMatrix(intRowCtr, 0) = Mid(rsrequest.Fields(intRowCtr).Name, 8)
'                .TextMatrix(intRowCtr, 6) = Mid(rsrequest.Fields(intRowCtr).Name, 8)
'                .TextMatrix(intRowCtr, 9) = rsrequest.Fields(intRowCtr).Name
'            Else
'                .TextMatrix(intRowCtr, 0) = rsrequest.Fields(intRowCtr).Name
'                .TextMatrix(intRowCtr, 6) = rsrequest.Fields(intRowCtr).Name
'                .TextMatrix(intRowCtr, 9) = rsrequest.Fields(intRowCtr).Name
'            End If
'        Next
'        '0 = Label
'        '1 = Value
'        '2 = Enable
'        '3 = Visible
'        '4 = Mandatory
'        '5 = Type
'        '6 = OrigFieldName
'        '7 = startrownew
'        '8 = startrowedt
'        '9 = DBField
'        '10 = optionalvalue
'        For intRowCtr = 0 To .Rows - 1
'            .Row = intRowCtr
'            .Col = 1
'            .CellAlignment = 0
'
'            .TextMatrix(.Row, 2) = getGlobalSettings("ENA_" & UCase(.TextMatrix(.Row, 0)), "RE_REQ", "0")
'            .TextMatrix(.Row, 3) = getGlobalSettings("VIS_" & UCase(.TextMatrix(.Row, 0)), "RE_REQ", "1")
'            .TextMatrix(.Row, 4) = getGlobalSettings("MAN_" & UCase(.TextMatrix(.Row, 0)), "RE_REQ", "1")
'            .TextMatrix(.Row, 5) = getGlobalSettings("TYPE_" & UCase(.TextMatrix(.Row, 0)), "RE_REQ", "TEXT")
'            .TextMatrix(.Row, 5) = IIf(.TextMatrix(.Row, 5) = "", "TEXT", .TextMatrix(.Row, 5))
'            .TextMatrix(.Row, 7) = IIf(.TextMatrix(.Row, 0) = strStartRowNewReq, strStartRowNewReq, "")
'
'            .CellBackColor = IIf(getGlobalSettings("ENA_" & UCase(.TextMatrix(.Row, 0)), "RE_REQ", "0") = "1", &H80000005, &H80000011)
'            .RowHeight(.Row) = IIf(.TextMatrix(.Row, 3) = "1", 350, 0)
'
'            .TextMatrix(.Row, 0) = getGlobalSettings(UCase(.TextMatrix(.Row, 0)), "RE_REQ", .TextMatrix(.Row, 0))
'            .TextMatrix(.Row, 8) = IIf(.TextMatrix(.Row, 0) = strStartRowEdtReq, strStartRowEdtReq, "")
'
'            If blnWithSettingsReqDet = False And .TextMatrix(.Row, 2) = "1" Then
'                blnWithSettingsReqDet = True
'            End If
'        Next
'
'        For intColCtr = 0 To .Cols - 1
'            .ColWidth(intColCtr) = 0
'        Next
'
'        .Row = 0
'        .ColWidth(0) = 1500
'        .ColWidth(1) = 2750
'
'        If blnNewPatient Then
'            .Enabled = False
'            For intRowCtr = 0 To .Rows - 1
'                .Row = intRowCtr
'                .CellBackColor = &H80000011
'            Next
'        Else
'            .Enabled = True
'        End If
'
'    End With
'End Sub
'
'Private Sub setmsfgExam(Optional strRequestID As String = "NULL")
'
'    With msfgExam
'        .Visible = False
'        .BackColor = &H80000011
'        strSQLStatement = IIf((strRequestID = "NULL"), "EXEC sp_tbLabRequestValues_select NULL", _
'                                "EXEC sp_tbLabRequestValues_select '" & strRequestID & "'")
'        Set rsrequest = myConnection.Execute(strSQLStatement)
'
'        If rsrequest.RecordCount > 0 Then
'            Set .DataSource = rsrequest
'            .Refresh
'            .Rows = .Rows + 1
'            .Enabled = True
'        Else
'            .Cols = rsrequest.Fields.Count + 1
'
'            For intColCtr = 1 To .Cols - 1
'                .TextArray(intColCtr) = rsrequest.Fields(intColCtr - 1).Name
'            Next
'            .Enabled = False
'        End If
'
'        For intRowCtr = 1 To .Rows - 1
'            .Row = intRowCtr
'            .Col = 0
'            .CellAlignment = 0
'            .CellBackColor = vbWhite
'
'            Select Case .TextMatrix(.Row, 5)
'                Case ""
'                    Set .CellPicture = imglstIcon.ListImages(6).Picture
'                Case "0"
'                    Set .CellPicture = imglstIcon.ListImages(6).Picture
'                Case "1"
'                    Set .CellPicture = imglstIcon.ListImages(2).Picture
'                Case "2"
'                    Set .CellPicture = imglstIcon.ListImages(3).Picture
'                Case "3"
'                    Set .CellPicture = imglstIcon.ListImages(4).Picture
'                Case "4"
'                    Set .CellPicture = imglstIcon.ListImages(8).Picture
'                Case "5"
'                    Set .CellPicture = imglstIcon.ListImages(7).Picture
'                Case "6"
'                    Set .CellPicture = imglstIcon.ListImages(9).Picture
'            End Select
'
'            .CellPictureAlignment = 4
'        Next
'
'        For intRowCtr = 1 To .Rows - 1
'            .RowHeight(intRowCtr) = 350
'        Next
'
'        .TextArray(0) = "Status"
'        .TextArray(1) = "Exam Code"
'        .TextArray(2) = "Section"
'        .TextArray(3) = "Request ID"
'        .TextArray(4) = "Section ID"
'        .TextArray(5) = "Status"
'        .TextArray(6) = "Sample ID"
'        .TextArray(7) = "Specimen"
'        .TextArray(8) = "Exam Name"
'
'        For intRowCtr = 0 To .Cols - 1
'            .ColWidth(intRowCtr) = 0
'        Next
'
'        .ColWidth(0) = 600
'        .ColWidth(1) = 1000
'        .ColWidth(2) = 2000
'        .ColWidth(6) = 2000
'        .ColWidth(7) = 1500
'        .ColWidth(8) = 4170
'
'        If rsrequest.RecordCount > 0 Then
'            For intRowCtr = 1 To .Rows - 1
'                AddExamDet intRowCtr
'            Next
'        End If
'        .Visible = True
'    End With
'
'End Sub
'
'Private Function blnExamOccured(examCriteria As String) As Boolean
'    Dim tempSectionID As String
'    Dim tempExamCode As String
'    Dim tempExamType As String
'    Dim tempSpecimen As String
'    Dim rsCheckParam As New ADODB.Recordset
'    Dim msfgExamDetRowCtr As Integer
'    Dim rsCheckParamRowCtr As Integer
'
'    tempExamType = getValue(4, examCriteria)
'    tempSectionID = getValue(5, examCriteria)
'    tempExamCode = getValue(1, examCriteria)
'    tempSpecimen = getValue(3, examCriteria)
'
'    strSQLStatement = "EXEC sp_getParamProfileExam_Select '" & tempExamType & "','" & tempSectionID & "'," & _
'                        "'" & tempExamCode & "','" & tempSpecimen & "'"
'
'    Set rsCheckParam = myConnection.Execute(strSQLStatement)
'
'    With msfgExamDet
'        For rsCheckParamRowCtr = 0 To rsCheckParam.RecordCount - 1
'            For msfgExamDetRowCtr = 1 To .Rows - 1
'                If RTrim(rsCheckParam.Fields(0).Value) = RTrim(.TextMatrix(msfgExamDetRowCtr, 0)) And _
'                    RTrim(rsCheckParam.Fields(2).Value) = RTrim(.TextMatrix(msfgExamDetRowCtr, 2)) And _
'                    RTrim(rsCheckParam.Fields(3).Value) = RTrim(.TextMatrix(msfgExamDetRowCtr, 3)) And _
'                    RTrim(rsCheckParam.Fields(4).Value) = RTrim(.TextMatrix(msfgExamDetRowCtr, 4)) Then
'                    blnExamOccured = True
'                    Exit For
'                End If
'            Next
'            If blnExamOccured Then
'                Exit For
'            Else
'                rsCheckParam.MoveNext
'            End If
'        Next
'    End With
'
''    Select Case examType
''        Case "P"
''
''
''            Set rsCheckParam = myConnection.Execute(strSQLStatement)
''
''            If rsCheckParam.RecordCount > 0 Then
''                With msfgExamDet
''                    If Len(.TextMatrix(1, 0)) > 0 Then
''                        For rsCheckParamRowCtr = 0 To rsCheckParam.RecordCount - 1
''                            For msfgExamDetRowCtr = 1 To .Rows - 1
''                                If RTrim(rsCheckParam.Fields(0).Value) = RTrim(.TextMatrix(msfgExamDetRowCtr, 0)) And _
''                                    RTrim(rsCheckParam.Fields(11).Value) = RTrim(.TextMatrix(msfgExamDetRowCtr, 1)) Then
''                                    blnExamOccured = True
''                                    Exit For
''                                End If
''                            Next
''                            If blnExamOccured Then
''                                Exit For
''                            Else
''                                rsCheckParam.MoveNext
''                            End If
''                        Next
''                    Else
''                        blnExamOccured = False
''                    End If
''                    rsCheckParam.Close
''                End With
''            Else
''                blnExamOccured = False
''            End If
''        Case "T"
''            With msfgExamDet
''                If Len(.TextMatrix(1, 0)) > 0 Then
''                    tempExamCode = getValue(6, examCriteria)
''                    For msfgExamDetRowCtr = 1 To .Rows - 1
''                        If tempInsID = RTrim(.TextMatrix(msfgExamDetRowCtr, 0)) And _
''                            tempExamCode = RTrim(.TextMatrix(msfgExamDetRowCtr, 1)) Then
''                            blnExamOccured = True
''                            Exit For
''                        End If
''                    Next
''                Else
''                    blnExamOccured = False
''                End If
''            End With
''    End Select
''
'
'End Function
'
'Private Sub AddExamDet(introw As Integer)
'
'    Dim rsExamDetails As New ADODB.Recordset
'    Dim RowCtr As Integer
'
'    With msfgExam
'
'        strSQLStatement = "EXEC sp_getParamProfileExam_Select '" & .TextMatrix(introw, 10) & "'," & _
'                                    "'" & .TextMatrix(introw, 4) & "'," & _
'                                    "'" & .TextMatrix(introw, 1) & "'," & _
'                                    "'" & .TextMatrix(introw, 7) & "'"
'
'        Set rsExamDetails = myConnection.Execute(strSQLStatement)
'
'        If rsExamDetails.RecordCount > 0 Then
'            For RowCtr = 0 To rsExamDetails.RecordCount - 1
'                msfgExamDet.TextMatrix(msfgExamDet.Rows - 1, 0) = rsExamDetails.Fields(0)
'                msfgExamDet.TextMatrix(msfgExamDet.Rows - 1, 1) = rsExamDetails.Fields(1)
'                msfgExamDet.TextMatrix(msfgExamDet.Rows - 1, 2) = rsExamDetails.Fields(2)
'                msfgExamDet.TextMatrix(msfgExamDet.Rows - 1, 3) = rsExamDetails.Fields(3)
'                msfgExamDet.TextMatrix(msfgExamDet.Rows - 1, 4) = rsExamDetails.Fields(4)
'                rsExamDetails.MoveNext
'                msfgExamDet.Rows = msfgExamDet.Rows + 1
'            Next
'        End If
'
'    End With
'
'End Sub
'
'Private Sub DelExamDet(reqExamCode As Integer)
'
'    Dim delRowCtr As Integer
'
'    With msfgExamDet
'        delRowCtr = 1
'        Do
'            If .TextMatrix(delRowCtr, 1) = reqExamCode Then
'                .RemoveItem (delRowCtr)
'            Else
'                delRowCtr = delRowCtr + 1
'            End If
'        Loop Until delRowCtr = .Rows - 1
'    End With
'
'End Sub
'
'Private Function ExtDemoConn() As String
'    Dim strDBType As String
'    Dim strServerName As String
'    Dim strDatabaseName As String
'    Dim strDBUserName As String
'    Dim strDBPassword As String
'
'    strDBType = GetSetting("HostForms_RE", "ExtDemoDB\Settings", "DBType", "")
'
'    Select Case strDBType
'        Case "MS ACCESS"
'        Case "MS SQL SERVER"
'            strServerName = Blowfish.DecryptString(GetSetting("HostForms_RE", "ExtDemoDB\Settings", "ServerName"), myKey, True)
'            strDatabaseName = Blowfish.DecryptString(GetSetting("HostForms_RE", "ExtDemoDB\Settings", "Database"), myKey, True)
'            strDBUserName = Blowfish.DecryptString(GetSetting("HostForms_RE", "ExtDemoDB\Settings", "UserName"), myKey, True)
'            strDBPassword = Blowfish.DecryptString(GetSetting("HostForms_RE", "ExtDemoDB\Settings", "Password"), myKey, True)
'
'
'            ExtDemoConn = "Provider=sqloledb;Data Source=" & strServerName & _
'                                ";Initial Catalog=" & strDatabaseName & _
'                                ";User Id=" & strDBUserName & _
'                                ";Password=" & strDBPassword
'        Case "MY SQL"
'        Case "ODBC"
'        Case Else
'            ExtDemoConn = "UNKNOWN"
'    End Select
'End Function
'
'Private Sub SetDemoFromExtDb(tempRs As ADODB.Recordset)
'    Dim myCtr As Integer
'
'    With msfgResult
'        If tempRs.RecordCount > 0 Then
'            For myCtr = 0 To .Rows - 1
'                .TextMatrix(myCtr, 1) = IIf(IsNull(tempRs.Fields(Trim(GetSetting("HostForms_RE", "ExtDemoDB\Mapping", _
'                                        .TextMatrix(myCtr, 0)))).Value), "", tempRs.Fields(Trim(GetSetting("HostForms_RE", "ExtDemoDB\Mapping", _
'                                        .TextMatrix(myCtr, 0)))).Value)
'
'            Next
'        Else
'            For myCtr = 0 To .Rows - 1
'                .TextMatrix(myCtr, 1) = ""
'            Next
'        End If
'    End With
'End Sub
'
'Private Sub sizeFix()
'    Me.Width = Screen.Width
'    With msfgRequest
'        .Width = .Width - 1400
'        .ColWidth(1) = .Width * 0.14
'    End With
'
'    cmdNewRequest.Width = cmdNewRequest.Width - 700
'    cmdCopyRequest.Width = cmdNewRequest.Width
'    'cmdDelReq.Width = cmdNewRequest.Width
'
'    cmdCopyRequest.Left = cmdNewRequest.Left + cmdNewRequest.Width + 30
'    'cmdDelReq.Left = cmdCopyRequest.Left + cmdCopyRequest.Width + 30
'
'    With msfgExam
'        .Width = msfgRequest.Width
'
'        .ColWidth(0) = 600
'        .ColWidth(1) = 1000
'        .ColWidth(2) = 2000
'        .ColWidth(6) = 2000
'        .ColWidth(7) = 1500
'        .ColWidth(8) = 4170
'
'    End With
'
'    cmdSaveExam.Width = cmdSaveExam.Width - 480
'    cmdPrintBarCode.Width = cmdSaveExam.Width
'    cmdDeleteExam.Width = cmdSaveExam.Width
'
'    cmdPrintBarCode.Left = cmdSaveExam.Left + cmdSaveExam.Width + 30
'    cmdDeleteExam.Left = cmdPrintBarCode.Left + cmdPrintBarCode.Width + 30
'
'    cmbSectionFilter.Left = cmdCopyRequest.Left + (cmbSectionFilter.Width / 2)
'    lblText(4).Left = cmbSectionFilter.Left - 800
'End Sub
'

