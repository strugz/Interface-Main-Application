VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmSearchResult 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search Result"
   ClientHeight    =   2955
   ClientLeft      =   4260
   ClientTop       =   4320
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picSend 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2190
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   20
      Top             =   540
      Visible         =   0   'False
      Width           =   405
      Begin MSComctlLib.Toolbar tbUpload 
         Height          =   330
         Left            =   30
         TabIndex        =   21
         Top             =   30
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlImages"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "send"
               Object.ToolTipText     =   "Upload Resutl to H.I.S."
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraPopUp 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   4860
      TabIndex        =   8
      Top             =   300
      Visible         =   0   'False
      Width           =   2355
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3300
         TabIndex        =   3
         Top             =   1200
         Width           =   1725
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Edit [Enter]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3300
         TabIndex        =   0
         Top             =   30
         Width           =   1725
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Preview [F5]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3300
         TabIndex        =   1
         Top             =   420
         Width           =   1725
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print [F6]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3300
         TabIndex        =   2
         Top             =   810
         Width           =   1725
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel [Esc]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3300
         TabIndex        =   4
         Top             =   1530
         Width           =   1725
      End
      Begin VB.Image Image8 
         Height          =   240
         Left            =   120
         Picture         =   "frmSearchResult.frx":0000
         Stretch         =   -1  'True
         Top             =   390
         Width           =   270
      End
      Begin VB.Label lblMerge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "           [F9] - Merge"
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
         Height          =   375
         Left            =   30
         TabIndex        =   17
         Top             =   360
         Width           =   2295
      End
      Begin VB.Image Image7 
         Height          =   240
         Left            =   150
         Picture         =   "frmSearchResult.frx":0442
         Top             =   750
         Width           =   240
      End
      Begin VB.Label lblUpload 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "           [F8] - Upload"
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
         Height          =   375
         Left            =   30
         TabIndex        =   16
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8B22&
         Caption         =   "lblSample"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   420
         TabIndex        =   15
         Top             =   4620
         Width           =   2895
      End
      Begin VB.Image Image6 
         Height          =   240
         Left            =   150
         Picture         =   "frmSearchResult.frx":09CC
         Top             =   2190
         Width           =   240
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   3330
         Picture         =   "frmSearchResult.frx":13CE
         Top             =   2070
         Width           =   240
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   150
         Picture         =   "frmSearchResult.frx":1DD0
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   150
         Picture         =   "frmSearchResult.frx":27D2
         Top             =   1440
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   120
         Picture         =   "frmSearchResult.frx":31D4
         Top             =   60
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   180
         Picture         =   "frmSearchResult.frx":3BD6
         Top             =   1110
         Width           =   240
      End
      Begin VB.Label lblArchive 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "          Archive / Delete"
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
         Height          =   375
         Left            =   3240
         TabIndex        =   14
         Top             =   2010
         Width           =   1755
      End
      Begin VB.Label lblCancel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "           Cancel Menu"
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
         Height          =   375
         Left            =   30
         TabIndex        =   13
         Top             =   2130
         Width           =   2295
      End
      Begin VB.Label lblDelete 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "           [Shift+ Del] - Delete"
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
         Height          =   375
         Left            =   30
         TabIndex        =   12
         Top             =   1770
         Width           =   2295
      End
      Begin VB.Label lblPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "           [F6] - Print"
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
         Height          =   375
         Left            =   30
         TabIndex        =   11
         Top             =   1410
         Width           =   2295
      End
      Begin VB.Label lblPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "           [F5] - Preview"
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
         Height          =   345
         Left            =   30
         TabIndex        =   10
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lblEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "          [Enter] - Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   2295
      End
   End
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7290
      TabIndex        =   6
      Top             =   3840
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Result"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   150
         Width           =   1350
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   4305
      Index           =   0
      Left            =   7620
      TabIndex        =   5
      Top             =   3210
      Visible         =   0   'False
      Width           =   7845
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
      Height          =   1785
      Left            =   90
      TabIndex        =   18
      Top             =   120
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   3149
      _Version        =   393216
      BackColor       =   16777215
      RowHeightMin    =   400
      BackColorFixed  =   10906411
      ForeColorFixed  =   16777215
      BackColorBkg    =   12632256
      GridColor       =   -2147483648
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
   Begin MSComctlLib.ImageList imlImages 
      Left            =   1260
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchResult.frx":4160
            Key             =   "saved"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchResult.frx":4B72
            Key             =   "printed"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchResult.frx":510C
            Key             =   "new"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchResult.frx":5B1E
            Key             =   "send"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   90
      TabIndex        =   19
      Top             =   2310
      Width           =   525
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   2340
      Top             =   2130
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   2700
      Top             =   2130
      Width           =   375
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   3060
      Top             =   2130
      Width           =   375
   End
End
Attribute VB_Name = "frmSearchResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cForms As New ClassForms
Private eConnection As New ADODB.Connection
Private rstExam2 As New ADODB.Recordset
Private mSortAsc As Boolean
Private yTrans As String

Private arrPatientName As String
Private arrResultDate As String
Private arrSequenceNum As String

Private Sub cmdCancel_Click()
    fraPopUp.Visible = False
    fraBody(0).Enabled = True
End Sub

Private Sub cmdDelete_Click()
Dim ctrStart As Integer
    ctrStart = IIf(msfgResult.Row > msfgResult.RowSel, msfgResult.RowSel, msfgResult.Row)
    fraPopUp.Visible = False
    fraBody(0).Enabled = True
    If msfgResult.Row = msfgResult.RowSel Then
        If MsgBox("Are you sure you want to delete this item?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
            DeleteResults
            setIcons msfgResult.TopRow
        Else
            msfgResult.SetFocus
        End If
    Else
        If MsgBox("Are you sure you want to delete these items?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
            DoEvents
            DeleteResults
            setIcons msfgResult.TopRow
        Else
            msfgResult.SetFocus
        End If
    End If
    If msfgResult.Rows <= ctrStart Then
        msfgResult.Row = msfgResult.Rows - 1
        msfgResult.ColSel = 3
    Else
        msfgResult.Row = ctrStart
        msfgResult.ColSel = 3
    End If
End Sub

Private Sub cmdOpen_Click()
    fraPopUp.Visible = False
    fraBody(0).Enabled = True
    msfgResult_DblClick
End Sub

Private Sub cmdPreview_Click()
Dim ctrStart As Long
Dim ctrLast As Long
    ctrStart = msfgResult.Row
    ctrLast = msfgResult.RowSel
    fraPopUp.Visible = False
    fraBody(0).Enabled = True
    If Len(Trim(msfgResult.TextMatrix(msfgResult.Row, 5))) > 0 Then
'        msfgResult.RowSel = msfgResult.Row
        Me.MousePointer = 11
        HostSA.showResultPrintPreview Trim(msfgResult.TextMatrix(msfgResult.Row, 5))
        Me.MousePointer = 0
        rstResult.Requery
        If rstResult.RecordCount = 0 Then
            frmSearchResult.Hide
        Else
            Set msfgResult.DataSource = rstResult
            lblTitle.Caption = "Search Result - [" & msfgResult.Rows - 1 & " Record(s) Found" & "]"
            msfgResult.SetFocus
            setIcons msfgResult.TopRow
            msfgResult.Row = ctrStart
            msfgResult.RowSel = ctrLast
            msfgResult.ColSel = 3
'            msfgResult.RowSel = msfgResult.Row
            msfgResult.Redraw = False
            msfgResult.Redraw = True
        End If
    Else
        If Val(GetSetting("Lab-I", "Settings", "Autoprocess", "0")) = 1 Then
            With msfgResult
                HostSA.showResultAutoPrint GetSetting("Lab-I", "Settings", "ProfileID", "1"), InstrumentID, .TextMatrix(.Row, 3), .TextMatrix(.Row, 6), .TextMatrix(.Row, 1)
            End With
        End If
    End If
End Sub

Private Sub cmdPrint_Click()
Dim ctrStart As Integer
Dim ctrLast As Integer
    ctrStart = msfgResult.Row
    ctrLast = msfgResult.RowSel
    fraPopUp.Visible = False
    fraBody(0).Enabled = True
    If MsgBox("Are you sure you want to Print the selected item(s)?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
        PrintResults
    End If
    msfgResult.Row = ctrStart
    msfgResult.RowSel = ctrLast
    msfgResult.ColSel = 3
End Sub

Private Sub Form_Deactivate()
    fraPopUp.Visible = False
    fraBody(0).Enabled = True
End Sub

Private Sub Form_Load()
setFormColors Me
Shape2.FillColor = GetSetting("Lab-I", "FormColors", "SavedDocCol", Shape2.FillColor)
Shape3.FillColor = GetSetting("Lab-I", "FormColors", "PrintedDocCol", Shape3.FillColor)
On Error GoTo errFound
    OpenLocalConnection
    frmSearchResult.Top = 1
    If mdiMain.Height > 800 Then frmSearchResult.Height = mdiMain.ScaleHeight  'mdiMain.Height - 800
    If mdiMain.Width > 5100 Then frmSearchResult.Width = mdiMain.ScaleWidth - frmSearch.Width   'mdiMain.Width - 5100
    If mdiMain.mnuSearch.Checked Or isConnected Then
        frmSearchResult.Left = frmSearch.Width  '4815 + 80
    Else
        frmSearchResult.Left = 1
    End If
    Set msfgResult.DataSource = rstResult
    Set rstExam2 = eConnection.Execute("Select * from tbcolabexam where active = 1")
    msfgResult.ColSel = 3
    lblStatus.Caption = ""
    lblTitle.Caption = "Search Result - [" & msfgResult.Rows - 1 & " Record(s) Found" & "] - Search Limit " & Val(GetSetting("Lab-I", "Settings", "SearchCount", "300"))
    mSortAsc = True
    With msfgResult
        setIcons .TopRow
        If .TextMatrix(.Row, 9) = 1 And .TextMatrix(.Row, 8) = "True" And Val(GetSetting("Lab-I", "VerSa", "EnablePlugin", "0")) = 1 Then
            picSend.Visible = True
            .Col = 3
            picSend.Top = .CellTop + .Top
            picSend.Left = .CellLeft + .Left + (.CellWidth - picSend.Width)
            .Col = 1
        Else
            picSend.Visible = False
        End If
        .ColSel = 3
    End With
errFound:
    If Err.Number = 30022 Then
        Unload Me
    ElseIf Err.Number = 0 Then
    Else
        MsgBox Err.Number & ": " & Err.Description, , App.ProductName
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
    msfgResult.RowHeightMin = 405
    fraTitle.Width = Me.Width - 10
    fraBody(0).Width = Me.Width - 10
    fraBody(0).Height = Me.Height - 300 '(fraTitle.Height - 180)
    msfgResult.Width = fraBody(0).Width - 250
    msfgResult.Height = Me.Height - 1000
    lblStatus.Top = Me.Height - 700
    With msfgResult
        .ColWidth(0) = .Width * 0.05 - (30 * 3)
        .ColWidth(1) = .Width * 0.3 - (30 * 3)
        .ColWidth(2) = .Width * 0.45 - (50 * 3)
        .ColWidth(3) = .Width * 0.2 - (20 * 3)
        .ColWidth(4) = 0
        .ColWidth(5) = 0
'        .ColWidth(4) = .Width * 0.2 - (20 * 3)
'        .ColWidth(5) = .Width * 0.2 - (20 * 3)
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .ColWidth(14) = 0
        .ColWidth(15) = 0
        If .Visible Then
            setIcons .TopRow
            If .TextMatrix(.Row, 9) = 1 And .TextMatrix(.Row, 8) = "True" And Val(GetSetting("Lab-I", "VerSa", "EnablePlugin", "0")) = 1 Then
                picSend.Visible = True
                .Col = 3
                picSend.Top = .CellTop + .Top
                picSend.Left = .CellLeft + .Left + (.CellWidth - picSend.Width)
                .Col = 1
                .ColSel = 3
            Else
                picSend.Visible = False
            End If
        End If
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    isResultVisible = False
'    eConnection.Close
End Sub

Private Sub Image1_Click()
lblPreview_Click
End Sub

Private Sub Image2_Click()
lblEdit_Click
End Sub

Private Sub Image3_Click()
lblPrint_Click
End Sub

Private Sub Image4_Click()
lblDelete_Click
End Sub

Private Sub Image5_Click()
lblArchive_Click
End Sub

Private Sub Image6_Click()
lblCancel_Click
End Sub

Private Sub Image7_Click()
lblUpload_Click
End Sub

Private Sub Image8_Click()
lblMerge_Click
End Sub

Private Sub lblArchive_Click()
'    fraPopUp.Visible = False
'    fraBody(0).Enabled = True
'    If msfgResult.Row = msfgResult.RowSel Then
'        If MsgBox("Are you sure you want to Archive this item?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
'            DeleteResults
'            setIcons msfgResult.Row
'        Else
'            msfgResult.SetFocus
'        End If
'    Else
'        If MsgBox("Are you sure you want to Archive these items?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
'            DeleteResults
'            setIcons msfgResult.Row
'        Else
'            msfgResult.SetFocus
'        End If
'    End If
End Sub

Private Sub lblArchive_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HighlightOver lblArchive
End Sub

Private Sub lblCancel_Click()
fraPopUp.Visible = False
cmdCancel_Click
End Sub

Private Sub lblCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HighlightOver lblCancel
End Sub

Private Sub lblDelete_Click()
fraPopUp.Visible = False
cmdDelete_Click
End Sub

Private Sub lblDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HighlightOver lblDelete
End Sub

Private Sub lblEdit_Click()
fraPopUp.Visible = False
cmdOpen_Click
End Sub

Private Sub lblEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HighlightOver lblEdit
End Sub

Private Sub lblMerge_Click()
Dim mRow As Integer
With msfgResult
    If CanMerge Then
        If .Row > .RowSel Then mRow = .RowSel Else mRow = .Row
        fraPopUp.Visible = False
        If MsgBox("Are you sure you want to merge results?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            MergeResults GetResultID
            rstResult.Requery
            Set .DataSource = rstResult
            setIcons .TopRow
            .Row = mRow
            .ColSel = 3
        End If
    End If
End With
End Sub

Private Sub lblMerge_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If CanMerge Then HighlightOver lblMerge Else HighlightOver lblMerge, vbButtonFace, vbGrayText
End Sub

Private Sub lblPreview_Click()
If lblPreview.BackColor <> vbButtonFace Then
    fraPopUp.Visible = False
    cmdPreview_Click
End If
End Sub

Private Sub lblPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If msfgResult.RowSel = msfgResult.Row Then HighlightOver lblPreview Else HighlightOver lblPreview, vbButtonFace, vbGrayText
End Sub

Private Sub lblPrint_Click()
fraPopUp.Visible = False
cmdPrint_Click
End Sub

Private Sub lblPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HighlightOver lblPrint
End Sub

Private Sub lblStatus_Click()
'Dim ctr, ctr1 As Integer
'Screen.MousePointer = 11
'For ctr = 1 To 100
'    myConnection.Execute "Insert into hr.labm_results (id,record_no,cs_no,exam_code,reslt_no,specimen_no) values (" & 1000 + ctr & "," & 1000 + ctr & "," & 1000 + ctr & "," & 1002 & "," & ctr & "," & 100 + ctr & ")"
''    For ctr1 = 1 To 13
''        If ctr1 >= 9 Then
''            myConnection.Execute "Insert into hr.labd_results (reslt_no, reslt_seq, reslt_code) values (" & ctr & "," & ctr1 & "," & ctr1 + 16 & ")"
''        Else
''            myConnection.Execute "Insert into hr.labd_results (reslt_no, reslt_seq, reslt_code) values (" & ctr & "," & ctr1 & "," & ctr1 + 13 & ")"
''        End If
''    Next ctr1
'Next ctr
'Screen.MousePointer = 0
End Sub

Private Sub lblTitle_Change()
    Me.Caption = lblTitle.Caption
End Sub

Private Sub lblUpload_Click()
Dim rstTemp As ADODB.Recordset
Dim tHeader As String
Dim tDetails As String
Dim tResultID As String
Dim ctr As Integer
If EnableDTP = 1 And msfgResult.RowSel = msfgResult.Row And Len(Trim(msfgResult.TextMatrix(msfgResult.Row, 5))) = 0 Then
    With msfgResult
    fraPopUp.Visible = False
    Set rstTemp = myConnection.Execute("Select * from tblabinstrumentresults where instrumentid = '" & InstrumentID & "' and PatientName = '" & .TextMatrix(.Row, 6) & "' and ResultDate = #" & .TextMatrix(.Row, 1) & "# and AccessionNumOrg = '" & .TextMatrix(.Row, 3) & "'")
    If rstTemp.RecordCount > 0 Then
        rstTemp.MoveFirst
        If Len(Trim(rstTemp.Fields("PatientName").Value & "")) > 0 Then
            tHeader = "|" & rstTemp.Fields("AccessionNumOrg").Value & "|" & rstTemp.Fields("NewPatientName").Value & "|" & rstTemp.Fields("ResultDate").Value
        Else
            tHeader = "|" & rstTemp.Fields("AccessionNumOrg").Value & "|" & rstTemp.Fields("PatientName").Value & "|" & rstTemp.Fields("ResultDate").Value
        End If
        tResultID = rstTemp.Fields("ResultID").Value
    End If
    Set rstTemp = myConnection.Execute("Select * from tblabinstrumentrValues where resultID ='" & tResultID & "'")
    If rstTemp.RecordCount > 0 Then
        rstTemp.MoveFirst
        tDetails = ""
        For ctr = 1 To rstTemp.RecordCount
            tDetails = tDetails & rstTemp("ExamName").Value & "|" & rstTemp("Result").Value & vbCrLf
            rstTemp.MoveNext
        Next ctr
    End If
    ObjDTP.UploadResults tHeader, tDetails
    Set rstTemp = Nothing
    End With
End If
End Sub

Private Sub lblUpload_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If EnableDTP = 1 And msfgResult.RowSel = msfgResult.Row And Len(Trim(msfgResult.TextMatrix(msfgResult.Row, 5))) = 0 Then HighlightOver lblUpload Else HighlightOver lblUpload, vbButtonFace, vbGrayText
End Sub

Private Sub msfgResult_Click()
If msfgResult.MouseRow = 0 Then
    If mSortAsc = True Then mSortAsc = False Else mSortAsc = True
    Select Case msfgResult.MouseCol
        Case 0
             Set rstResult = myConnection.Execute(frmSearch.getSearchString(0, mSortAsc))
             Set msfgResult.DataSource = rstResult
        Case 1
             Set rstResult = myConnection.Execute(frmSearch.getSearchString(1, mSortAsc))
             Set msfgResult.DataSource = rstResult
        Case 2
             Set rstResult = myConnection.Execute(frmSearch.getSearchString(2, mSortAsc))
             Set msfgResult.DataSource = rstResult
        Case 3
             Set rstResult = myConnection.Execute(frmSearch.getSearchString(3, mSortAsc))
             Set msfgResult.DataSource = rstResult
    End Select
    msfgResult.Row = 1
    msfgResult.ColSel = 3
    setIcons msfgResult.TopRow
End If
End Sub

Private Sub setSelectedItems()
Dim ctr As Integer
Dim StartRow As Integer
Dim lastRow As Integer
    arrPatientName = ""
    arrResultDate = ""
    arrSequenceNum = ""
    If msfgResult.RowSel <> msfgResult.Row And ResultEditingMethod = "B" Then
        If msfgResult.Row > msfgResult.RowSel Then
            StartRow = msfgResult.RowSel
            lastRow = msfgResult.Row
        Else
            StartRow = msfgResult.Row
            lastRow = msfgResult.RowSel
        End If
        For ctr = StartRow To lastRow
            If ctr <> lastRow Then
                arrPatientName = arrPatientName & msfgResult.TextMatrix(ctr, 6) & "|"
                arrResultDate = arrResultDate & msfgResult.TextMatrix(ctr, 1) & "|"
                arrSequenceNum = arrSequenceNum & msfgResult.TextMatrix(ctr, 3) & "|"
            Else
                arrPatientName = arrPatientName & msfgResult.TextMatrix(ctr, 6)
                arrResultDate = arrResultDate & msfgResult.TextMatrix(ctr, 1)
                arrSequenceNum = arrSequenceNum & msfgResult.TextMatrix(ctr, 3)
            End If
        Next ctr
    Else
        arrPatientName = arrPatientName & msfgResult.TextMatrix(msfgResult.Row, 6)
        arrResultDate = arrResultDate & msfgResult.TextMatrix(msfgResult.Row, 1)
        arrSequenceNum = arrSequenceNum & msfgResult.TextMatrix(msfgResult.Row, 3)
    End If
End Sub

Private Sub msfgResult_DblClick()
Dim rstTemp As ADODB.Recordset
Dim ctrStart As Long
Dim ctrLast As Long
Dim ctr As Long
Dim tempstr As String
Dim myArray() As String
Dim arrPName() As String
Dim arrRDate() As String
Dim arrSNum() As String
'MakeTransparent mdiMain.hWnd, 240
With msfgResult
.Redraw = False
If .MouseRow > 0 Then
    If DBServerType = "LOCAL DATABASE" Or DBServerType = "MSSQL SERVER" Then
        With msfgResult
            If .Row >= .RowSel Then
                ctrStart = .RowSel
                ctrLast = .Row
            Else
                ctrStart = .Row
                ctrLast = .RowSel
            End If
            setSelectedItems
            tempstr = HostSA.showResultEntry(InstrumentID, arrPatientName, arrResultDate, arrSequenceNum, ctrStart, ctrLast, ctrStart)
            myArray = Split(tempstr, "|")
        End With
    ElseIf DBServerType = "MSSQL SERVER" Then
        cForms.ShowResultViewer InstrumentID, .TextMatrix(.Row, 2), .TextMatrix(.Row, 1)
    ElseIf DBServerType = "MYSQL SERVER" Then
        With msfgResult
            If .Row >= .RowSel Then
                ctrStart = .RowSel
                ctrLast = .Row
            Else
                ctrStart = .Row
                ctrLast = .RowSel
            End If
            setSelectedItems
            tempstr = HostSA.showResultEntry(InstrumentID, arrPatientName, arrResultDate, arrSequenceNum, ctrStart, ctrLast, ctrStart)
            myArray = Split(tempstr, "|")
        End With
    Else
        cForms.ShowResultViewer InstrumentID, .TextMatrix(.Row, 2), .TextMatrix(.Row, 1)
    End If
    If rstResult.RecordCount = 0 Then
        frmSearchResult.Hide
    Else
        If Val(GetSetting("HostForms_SA", "Settings", "UpdateSearchResult", "0")) = 1 Then
            arrPName = Split(arrPatientName, "|")
            arrRDate = Split(arrResultDate, "|")
            arrSNum = Split(arrSequenceNum, "|")
            For ctr = ctrStart To ctrLast
                Set rstTemp = myConnection.Execute("Select * from tbLabInstrumentResults where InstrumentID = '" & InstrumentID & "' and ResultID = '" & .TextMatrix(ctr, 10) & "'")
                If rstTemp.RecordCount > 0 Then
                    .TextMatrix(ctr, 8) = rstTemp.Fields("printed").Value & ""
                    .TextMatrix(ctr, 9) = rstTemp.Fields("status").Value & ""
                    .TextMatrix(ctr, 7) = rstTemp.Fields("newpatientname").Value & ""
                    .TextMatrix(ctr, 3) = rstTemp.Fields("accessionnum").Value & ""
                    .TextMatrix(ctr, 5) = rstTemp.Fields("refnum").Value & ""
                End If
            Next ctr
            setIcons ctrStart, ctrLast
        Else
            setIcons .TopRow
        End If
        lblTitle.Caption = "Search Result - [" & msfgResult.Rows - 1 & " Record(s) Found" & "]"
        .SetFocus
        .Row = ctrStart
        .RowSel = ctrLast
        
        If .TextMatrix(.Row, 9) = 1 And .TextMatrix(.Row, 8) = "True" And Val(GetSetting("Lab-I", "VerSa", "EnablePlugin", "0")) = 1 Then
            picSend.Visible = True
            .Col = 3
            picSend.Top = .CellTop + .Top
            picSend.Left = .CellLeft + .Left + (.CellWidth - picSend.Width)
            .Col = 1
        Else
            picSend.Visible = False
        End If
        .ColSel = 3
    End If
End If
.Redraw = True
End With
Set rstTemp = Nothing
End Sub

Private Sub msfgResult_EnterCell()
    fraPopUp.Visible = False
    With msfgResult
        .Redraw = False
        If .Visible And .Enabled Then
            .SetFocus
            If .TextMatrix(.Row, 9) = 1 And .TextMatrix(.Row, 8) = "True" And Val(GetSetting("Lab-I", "VerSa", "EnablePlugin", "0")) = 1 Then
                picSend.Visible = True
                .Col = 3
                picSend.Top = .CellTop + .Top
                picSend.Left = .CellLeft + .Left + (.CellWidth - picSend.Width)
                .Col = 1
            Else
                picSend.Visible = False
            End If
        End If
        .Redraw = True
    End With
End Sub

Private Sub msfgResult_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tempno As String
If KeyCode = vbKeyReturn Then
    msfgResult_DblClick
End If
If KeyCode = vbKeyDelete And Shift = 1 Then
    If msfgResult.Row = msfgResult.RowSel Then
        If MsgBox("Are you sure you want to delete this item?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
            DeleteResults
            setIcons msfgResult.TopRow
        Else
            msfgResult.SetFocus
        End If
    Else
        If MsgBox("Are you sure you want to delete these items?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
            DeleteResults
            setIcons msfgResult.TopRow
        Else
            msfgResult.SetFocus
        End If
    End If
End If
'If KeyCode = vbKeyF1 Then
'    cForms.ShowResultViewer InstrumentID, msfgResult.TextMatrix(msfgResult.Row, 2), msfgResult.TextMatrix(msfgResult.Row, 1)
'End If
'If KeyCode = vbKeyF8 Then
'    If msfgResult.Row = msfgResult.RowSel Then
'        If MsgBox("Are you sure you want to transfer this item?", vbYesNo + vbQuestion + vbDefaultButton2, "Transfer Item") = vbYes Then
'            TransferResults
'        Else
'            msfgResult.SetFocus
'        End If
'    Else
'        If MsgBox("Are you sure you want to transfer these items?", vbYesNo + vbQuestion + vbDefaultButton2, "Transfer Items") = vbYes Then
'            TransferResults
'        Else
'            msfgResult.SetFocus
'        End If
'    End If
'End If
If KeyCode = vbKeyF6 Then
'    If MsgBox("Are you sure you want to Print the selected item(s)?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
        cmdPrint_Click
        'PrintResults
'    End If
End If
If KeyCode = vbKeyF2 Then
    MsgBox "Enter Starting Number for the Accession Number?", vbOKOnly + vbInformation, App.ProductName
    tempno = cForms.showInputBox("Batch Update - Accession Number", , , , 8, True)
    If IsNumeric(tempno) And Val(tempno) >= 0 And Len(tempno) > 0 Then
        UpdateAccNum tempno
    Else
        MsgBox "Invalid Accession Number!", vbInformation + vbOKOnly, App.ProductName
    End If
End If
If KeyCode = 8 Or KeyCode = vbKeyEscape Then
    Unload Me
End If
If KeyCode = vbKeyF5 Then
    If Len(Trim(msfgResult.TextMatrix(msfgResult.Row, 5))) > 0 Then
'        msfgResult.RowSel = msfgResult.Row
        Me.MousePointer = 11
        HostSA.showResultPrintPreview Trim(msfgResult.TextMatrix(msfgResult.Row, 5))
        Me.MousePointer = 0
        rstResult.Requery
        If rstResult.RecordCount = 0 Then
            frmSearchResult.Hide
        Else
            Set msfgResult.DataSource = rstResult
            lblTitle.Caption = "Search Result - [" & msfgResult.Rows - 1 & " Record(s) Found" & "]"
            msfgResult.SetFocus
            setIcons msfgResult.TopRow
            msfgResult.RowSel = msfgResult.Row
            msfgResult.Redraw = False
            msfgResult.Redraw = True
        End If
    End If
End If
If KeyCode = vbKeyEscape Then
    isResultVisible = False
    Unload frmSearchResult
End If
If KeyCode = vbKeyA And Shift = 1 Then
    msfgResult.Row = 1
    msfgResult.RowSel = msfgResult.Rows - 1
'    msfgResult.Col = 1
    msfgResult.ColSel = 6
End If
If KeyCode = vbKeyF8 Then
    lblUpload_Click
End If
If KeyCode = vbKeyF9 Then
    lblMerge_Click
End If
End Sub

Private Sub UpdateAccNum(myStartNum As String)
Dim ctr As Long
Dim temprow As Long
Dim tempnum As Long
Dim tempCount As Long
tempnum = Val(myStartNum)
tempCount = 0
Screen.MousePointer = 11
With msfgResult
    temprow = .Row
    If .Row <= .RowSel Then
        For ctr = .Row To .RowSel
            lblStatus.Caption = "Updating Record: " & .TextMatrix(ctr, 2)
            If DBServerType = "MSSQL SERVER" Then
                myConnection.Execute "Update tbLabInstrumentResults set Accessionnum = '" & tempnum + tempCount & "' where Instrumentid = '" & InstrumentID & "' and Patientname = '" & .TextMatrix(ctr, 2) & "' and convert(varchar(16),resultdate,120) = '" & Format(.TextMatrix(ctr, 1), "yyyy-mm-dd HH:mm") & "'"
            ElseIf DBServerType = "ORACLE" Then
                myConnection.Execute Replace("Update <schema>tbLabInstrumentResults set Accessionnum = '" & tempnum + tempCount & "' where Instrumentid = '" & InstrumentID & "' and Patientname = '" & .TextMatrix(ctr, 2) & "' and to_char(ResultDate,'MM/DD/YYYY HH24:MI') = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "'", "<schema>", DBSchema)
            Else
                If myStartNum = 0 Then
                    myConnection.Execute "Update tbLabInstrumentResults set Accessionnum = '' where Instrumentid = '" & InstrumentID & "' and Patientname = '" & .TextMatrix(ctr, 2) & "' and format(resultdate,""yyyy-mm-dd HH:mm"") = '" & Format(.TextMatrix(ctr, 1), "yyyy-mm-dd HH:mm") & "'"
                Else
                    myConnection.Execute "Update tbLabInstrumentResults set Accessionnum = '" & tempnum + tempCount & "' where Instrumentid = '" & InstrumentID & "' and Patientname = '" & .TextMatrix(ctr, 2) & "' and format(resultdate,""yyyy-mm-dd HH:mm"") = '" & Format(.TextMatrix(ctr, 1), "yyyy-mm-dd HH:mm") & "'"
                End If
            End If
            tempCount = tempCount + 1
        Next ctr
    Else
        For ctr = .RowSel To .Row
            lblStatus.Caption = "Updating Record: " & .TextMatrix(ctr, 2)
            If DBServerType = "MSSQL SERVER" Then
                myConnection.Execute "Update tbLabInstrumentResults set Accessionnum = '" & tempnum + tempCount & "' where Instrumentid = '" & InstrumentID & "' and Patientname = '" & .TextMatrix(ctr, 2) & "' and convert(varchar(16),resultdate,120) = '" & Format(.TextMatrix(ctr, 1), "yyyy-mm-dd HH:mm") & "'"
            ElseIf DBServerType = "ORACLE" Then
                myConnection.Execute Replace("Update <schema>tbLabInstrumentResults set Accessionnum = '" & tempnum + tempCount & "' where Instrumentid = '" & InstrumentID & "' and Patientname = '" & .TextMatrix(ctr, 2) & "' and to_char(ResultDate,'MM/DD/YYYY HH24:MI') = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "'", "<schema>", DBSchema)
            Else
                If myStartNum = 0 Then
                    myConnection.Execute "Update tbLabInstrumentResults set Accessionnum = '' where Instrumentid = '" & InstrumentID & "' and Patientname = '" & .TextMatrix(ctr, 2) & "' and format(resultdate,""yyyy-mm-dd HH:mm"") = '" & Format(.TextMatrix(ctr, 1), "yyyy-mm-dd HH:mm") & "'"
                Else
                    myConnection.Execute "Update tbLabInstrumentResults set Accessionnum = '" & tempnum + tempCount & "' where Instrumentid = '" & InstrumentID & "' and Patientname = '" & .TextMatrix(ctr, 2) & "' and format(resultdate,""yyyy-mm-dd HH:mm"") = '" & Format(.TextMatrix(ctr, 1), "yyyy-mm-dd HH:mm") & "'"
                End If
            End If
            tempCount = tempCount + 1
        Next ctr
    End If
lblStatus.Caption = ""
rstResult.Requery
If rstResult.RecordCount = 0 Then
    frmSearchResult.Hide
Else
    Set msfgResult.DataSource = rstResult
    If temprow > .Rows - 1 Then .Row = .Rows - 1 Else .Row = temprow
    .ColSel = 3
    lblTitle.Caption = "Search Result - [" & msfgResult.Rows - 1 & " Record(s) Found" & "]"
    msfgResult.SetFocus
End If
End With
Screen.MousePointer = 0
End Sub

Private Sub DeleteResults()
Dim ctr As Integer
Dim temprow As Integer
Screen.MousePointer = 11
myConnection.BeginTrans
With msfgResult
    temprow = .Row
    If .Row <= .RowSel Then
        For ctr = .Row To .RowSel
            If Len(Trim(.TextMatrix(ctr, 5) & "")) > 0 Then
                myConnection.Execute "Delete from tblabinstrumentRValues where resultid = (select distinct top 1 resultid from tblabinstrumentResults where refnum = '" & .TextMatrix(ctr, 5) & "')"
                myConnection.Execute "Delete from tbPatientValues where refnum = '" & .TextMatrix(ctr, 5) & "'"
                myConnection.Execute "Delete from tbPatientMaster where refnum = '" & .TextMatrix(ctr, 5) & "'"
                myConnection.Execute "Delete from tbLabInstrumentResults where refnum = '" & .TextMatrix(ctr, 5) & "'"
            Else
                lblStatus.Caption = "Deleting Record: " & .TextMatrix(ctr, 2)
                If DBServerType = "MSSQL SERVER" Then
                    myConnection.Execute "Delete tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum = '" & .TextMatrix(ctr, 3) & "' and convert(varchar(10),ResultDate,101) = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy") & "'"
                ElseIf DBServerType = "ORACLE" Then
                    If Len(.TextMatrix(ctr, 3)) > 0 Then
                        myConnection.Execute Replace("Delete from <schema>tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum = '" & .TextMatrix(ctr, 3) & "' and to_char(ResultDate,'MM/DD/YYYY HH24:MI') = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "'", "<schema>", DBSchema)
                    Else
                        myConnection.Execute Replace("Delete from <schema>tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum is null and to_char(ResultDate,'MM/DD/YYYY HH24:MI') = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "'", "<schema>", DBSchema)
                    End If
                Else
                    If Trim(Len(.TextMatrix(.Row, 5))) < 1 Then
                        myConnection.Execute "Delete from tblabInstrumentRValues where Resultid = (Select distinct top 1 Resultid from tblabinstrumentresults where PatientName = '" & Replace(.TextMatrix(ctr, 6), "'", "''") & "' and Accessionnum = '" & Replace(.TextMatrix(ctr, 3), "'", "''") & "' and format(ResultDate,""mm/dd/yyyy HH:mm"") = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "')"
                        myConnection.Execute "Delete from tblabinstrumentresults where PatientName = '" & Replace(.TextMatrix(ctr, 6), "'", "''") & "' and Accessionnum = '" & Replace(.TextMatrix(ctr, 3), "'", "''") & "' and format(ResultDate,""mm/dd/yyyy HH:mm"") = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "'"
                    End If
                End If
            End If
        Next ctr
    Else
        For ctr = .RowSel To .Row
            lblStatus.Caption = "Deleting Record: " & .TextMatrix(ctr, 2)
            If DBServerType = "MSSQL SERVER" Then
                myConnection.Execute "Delete tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum = '" & .TextMatrix(ctr, 3) & "' and convert(varchar(10),ResultDate,101) = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy") & "'"
            ElseIf DBServerType = "ORACLE" Then
                If Len(.TextMatrix(ctr, 3)) > 0 Then
                    myConnection.Execute Replace("Delete from <schema>tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum = '" & .TextMatrix(ctr, 3) & "' and to_char(ResultDate,'MM/DD/YYYY HH24:MI') = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "'", "<schema>", DBSchema)
                Else
                    myConnection.Execute Replace("Delete from <schema>tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum is null and to_char(ResultDate,'MM/DD/YYYY HH24:MI') = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "'", "<schema>", DBSchema)
                End If
            Else
                myConnection.Execute "Delete from tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum = '" & .TextMatrix(ctr, 3) & "' and format(ResultDate,""mm/dd/yyyy HH:mm"") = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "'"
            End If
'            myConnection.Execute "Delete tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum = '" & .TextMatrix(ctr, 3) & "' and convert(varchar(10),ResultDate,101) = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy") & "'"
            myConnection.Execute "Delete from tblabinstrumentresults where PatientName = '" & Replace(.TextMatrix(ctr, 2), "'", "''") & "' and Accessionnum = '" & Replace(.TextMatrix(ctr, 3), "'", "''") & "' and format(ResultDate,""mm/dd/yyyy HH:mm"") = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "'"
        Next ctr
    End If
lblStatus.Caption = ""
rstResult.Requery
If rstResult.RecordCount = 0 Then
    frmSearchResult.Hide
Else
    Set msfgResult.DataSource = rstResult
    If temprow > .Rows - 1 Then .Row = .Rows - 1 Else .Row = temprow
    .ColSel = 3
    lblTitle.Caption = "Search Result - [" & msfgResult.Rows - 1 & " Record(s) Found" & "]"
    msfgResult.SetFocus
End If
End With
myConnection.CommitTrans
Screen.MousePointer = 0
End Sub

Private Sub ArchiveResults()
Dim ctr As Integer
Dim temprow As Integer
Screen.MousePointer = 11
With msfgResult
    temprow = .Row
    If .Row <= .RowSel Then
        For ctr = .Row To .RowSel
            If Len(Trim(.TextMatrix(ctr, 5) & "")) > 0 Then
                If MsgBox("Result of Patient: " & .TextMatrix(ctr, 2) & " has already been saved." & vbCrLf & _
                    "Are you sure you want to delete this record?", vbQuestion + vbYesNo + vbDefaultButton2, App.ProductName) = vbYes Then
                    myConnection.Execute "Delete from tbLabInstrumentResults where refnum = '" & .TextMatrix(ctr, 5) & "'"
                End If
            Else
                lblStatus.Caption = "Deleting Record: " & .TextMatrix(ctr, 2)
                If DBServerType = "MSSQL SERVER" Then
                    myConnection.Execute "Delete tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum = '" & .TextMatrix(ctr, 3) & "' and convert(varchar(10),ResultDate,101) = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy") & "'"
                ElseIf DBServerType = "ORACLE" Then
                    If Len(.TextMatrix(ctr, 3)) > 0 Then
                        myConnection.Execute Replace("Delete from <schema>tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum = '" & .TextMatrix(ctr, 3) & "' and to_char(ResultDate,'MM/DD/YYYY HH24:MI') = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "'", "<schema>", DBSchema)
                    Else
                        myConnection.Execute Replace("Delete from <schema>tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum is null and to_char(ResultDate,'MM/DD/YYYY HH24:MI') = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "'", "<schema>", DBSchema)
                    End If
                Else
                    If Trim(Len(.TextMatrix(.Row, 5))) < 1 Then
                        myConnection.Execute "Delete from tblabinstrumentresults where PatientName = '" & Replace(.TextMatrix(ctr, 2), "'", "''") & "' and Accessionnum = '" & Replace(.TextMatrix(ctr, 3), "'", "''") & "' and format(ResultDate,""mm/dd/yyyy HH:mm"") = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "'"
                    End If
                End If
            End If
        Next ctr
    Else
        For ctr = .RowSel To .Row
            lblStatus.Caption = "Deleting Record: " & .TextMatrix(ctr, 2)
            If DBServerType = "MSSQL SERVER" Then
                myConnection.Execute "Delete tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum = '" & .TextMatrix(ctr, 3) & "' and convert(varchar(10),ResultDate,101) = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy") & "'"
            ElseIf DBServerType = "ORACLE" Then
                If Len(.TextMatrix(ctr, 3)) > 0 Then
                    myConnection.Execute Replace("Delete from <schema>tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum = '" & .TextMatrix(ctr, 3) & "' and to_char(ResultDate,'MM/DD/YYYY HH24:MI') = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "'", "<schema>", DBSchema)
                Else
                    myConnection.Execute Replace("Delete from <schema>tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum is null and to_char(ResultDate,'MM/DD/YYYY HH24:MI') = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "'", "<schema>", DBSchema)
                End If
            Else
                myConnection.Execute "Delete from tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum = '" & .TextMatrix(ctr, 3) & "' and format(ResultDate,""mm/dd/yyyy HH:mm"") = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "'"
            End If
'            myConnection.Execute "Delete tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum = '" & .TextMatrix(ctr, 3) & "' and convert(varchar(10),ResultDate,101) = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy") & "'"
            myConnection.Execute "Delete from tblabinstrumentresults where PatientName = '" & Replace(.TextMatrix(ctr, 2), "'", "''") & "' and Accessionnum = '" & Replace(.TextMatrix(ctr, 3), "'", "''") & "' and format(ResultDate,""mm/dd/yyyy HH:mm"") = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "'"
        Next ctr
    End If
lblStatus.Caption = ""
rstResult.Requery
If rstResult.RecordCount = 0 Then
    frmSearchResult.Hide
Else
    Set msfgResult.DataSource = rstResult
    If temprow > .Rows - 1 Then .Row = .Rows - 1 Else .Row = temprow
    .ColSel = 3
    lblTitle.Caption = "Search Result - [" & msfgResult.Rows - 1 & " Record(s) Found" & "]"
    msfgResult.SetFocus
End If
End With
Screen.MousePointer = 0
End Sub

Private Sub msfgResult_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbKeyRButton Then
    msfgResult.Redraw = False
    fraPopUp.Visible = False
    If msfgResult.Row = msfgResult.RowSel Then
        msfgResult.Row = msfgResult.MouseRow
        msfgResult.ColSel = 3
    End If
    msfgResult_Click
    msfgResult.Redraw = True
End If
End Sub

Private Sub msfgResult_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
msfgResult.Redraw = False
If Button = vbKeyRButton Then
    If X + 100 + fraPopUp.Width > msfgResult.Width Then
        fraPopUp.Left = X + 100 - fraPopUp.Width
    Else
        fraPopUp.Left = X + 100
    End If
    If Y + 100 + fraPopUp.Height > msfgResult.Height Then
        fraPopUp.Top = Y + 500 - fraPopUp.Height
    Else
        fraPopUp.Top = Y + 500
    End If
    fraPopUp.Visible = True
Else
    fraPopUp.Visible = False
End If
msfgResult.Redraw = True
End Sub

Private Sub msfgResult_Scroll()
Dim temprow As Long
fraPopUp.Visible = False
picSend.Visible = False
    temprow = msfgResult.Row
    setIcons msfgResult.TopRow
    msfgResult.Row = temprow
    msfgResult.ColSel = 3
End Sub

Private Sub msfgResult_SelChange()
    fraPopUp.Visible = False
    msfgResult.SetFocus
    If msfgResult.Row <= msfgResult.RowSel Then
        lblStatus.Caption = (msfgResult.RowSel - msfgResult.Row) + 1 & " Record(s) selected."
    Else
        lblStatus.Caption = (msfgResult.Row - msfgResult.RowSel) + 1 & " Record(s) selected."
    End If
    If msfgResult.Row <> msfgResult.RowSel Then picSend.Visible = False
End Sub

Private Sub TransferResults()
Dim temps As String
Dim ctr As Integer
Dim temprow As Integer
Dim rstDump As ADODB.Recordset
Dim rstCount As ADODB.Recordset
Dim rstExams As ADODB.Recordset
Dim rstSearch As ADODB.Recordset
Screen.MousePointer = 11
Set rstExams = eConnection.Execute("qrySelectActiveExam")
With msfgResult
    temprow = .Row
    rstExams.MoveFirst
    If rstExams.RecordCount > 0 Then
        If .Row <= .RowSel Then
            For ctr = .Row To .RowSel
                lblStatus.Caption = "Transfering Record: " & .TextMatrix(ctr, 2)
                If DBServerType = "MSSQL SERVER" Then
                    If Len(Trim(.TextMatrix(ctr, 3))) > 0 Then
                        Do While Not rstExams.EOF
                            Set rstCount = myConnection.Execute("select reslt_no from labm_results where specimen_no = " & .TextMatrix(ctr, 3) & " and convert(varchar(10),reslt_date,101) = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy") & "'")
                            MsgBox rstCount.RecordCount
                            rstExams.MoveNext
                        Loop
                    End If
                ElseIf DBServerType = "ORACLE" Then
                    If Len(Trim(.TextMatrix(ctr, 3))) > 0 Then
                        Do While Not rstExams.EOF
                            Set rstCount = myConnection.Execute(Replace("select reslt_no from <schema>labm_results where specimen_no = " & .TextMatrix(ctr, 3) & " and to_char(date_received,'mm/dd/yyyy') = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy") & "'", "<schema>", DBSchema))
                            If rstCount.RecordCount = 1 Then
                                temps = "Select * from <schema>tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum = '" & .TextMatrix(ctr, 3) & "' and to_char(ResultDate,'MM/DD/YYYY HH24:MI') = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "' and EXAMNAME = '" & rstExams.Fields("InstExamName").Value & "'"
                                Set rstDump = myConnection.Execute(Replace(temps, "<schema>", DBSchema))
                                
                                Set rstSearch = myConnection.Execute(Replace("Select * from <schema>LABM_RESULTS where EXAM_CODE = '" & rstExams.Fields("HisExamCode").Value & "' and SPECIMEN_NO = '" & .TextMatrix(ctr, 3) & "' and TO_CHAR(DATE_RECEIVED,'MM/DD/YYYY') = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy") & "'", "<schema>", DBSchema))
                                If rstDump.RecordCount = 1 And rstSearch.RecordCount = 1 Then
                                    temps = "Update <schema>LABD_RESULTS set RESLT_VALUE = " & rstDump.Fields("RESULT").Value & " where RESLT_NO = " & rstSearch.Fields("RESLT_NO").Value & " and RESLT_CODE = " & rstExams.Fields("HisResultCode").Value & " and RESLT_VALUE is null"
                                    myConnection.Execute Replace(temps, "<schema>", DBSchema)
                                    
                                    temps = "Delete from <schema>tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum = '" & .TextMatrix(ctr, 3) & "' and to_char(ResultDate,'MM/DD/YYYY HH24:MI') = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "' and EXAMNAME = '" & rstExams.Fields("InstExamName").Value & "'"
                                    myConnection.Execute Replace(temps, "<schema>", DBSchema)
                                End If
                            End If
                            rstExams.MoveNext
                        Loop
                        rstExams.MoveFirst
                    Else
                        MsgBox "Unable to transfer result. Click ok to continue", vbOKOnly + vbInformation, App.ProductName
                    End If
                Else
                    MsgBox "Unable to transfer result. Please check database settings. Click ok to continue", vbOKOnly + vbInformation, App.ProductName
                End If
            Next ctr
        Else
            For ctr = .RowSel To .Row
                lblStatus.Caption = "Transfering Record: " & .TextMatrix(ctr, 2)
                If DBServerType = "MSSQL SERVER" Then
    '                myConnection.Execute "Delete tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum = '" & .TextMatrix(ctr, 3) & "' and convert(varchar(10),ResultDate,101) = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy") & "'"
                ElseIf DBServerType = "ORACLE" Then
                    If Len(Trim(.TextMatrix(ctr, 3))) > 0 Then
                        Do While Not rstExams.EOF
                            Set rstCount = myConnection.Execute(Replace("select reslt_no from <schema>labm_results where specimen_no = " & .TextMatrix(ctr, 3) & " and to_char(date_received,'mm/dd/yyyy') = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy") & "'", "<schema>", DBSchema))
                            If rstCount.RecordCount = 1 Then
                                temps = "Select * from <schema>tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum = '" & .TextMatrix(ctr, 3) & "' and to_char(ResultDate,'MM/DD/YYYY HH24:MI') = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "' and EXAMNAME = '" & rstExams.Fields("InstExamName").Value & "'"
                                Set rstDump = myConnection.Execute(Replace(temps, "<schema>", DBSchema))
                                
                                Set rstSearch = myConnection.Execute(Replace("Select * from <schema>LABM_RESULTS where EXAM_CODE = '" & rstExams.Fields("HisExamCode").Value & "' and SPECIMEN_NO = '" & .TextMatrix(ctr, 3) & "' and TO_CHAR(DATE_RECEIVED,'MM/DD/YYYY') = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy") & "'", "<schema>", DBSchema))
                                If rstDump.RecordCount = 1 And rstSearch.RecordCount = 1 Then
                                    temps = "Update <schema>LABD_RESULTS set RESLT_VALUE = " & rstDump.Fields("RESULT").Value & " where RESLT_NO = " & rstSearch.Fields("RESLT_NO").Value & " and RESLT_CODE = " & rstExams.Fields("HisResultCode").Value & " and RESLT_VALUE is null"
                                    myConnection.Execute Replace(temps, "<schema>", DBSchema)
                                    
                                    temps = "Delete from <schema>tblabinstrumentresults where PatientName = '" & .TextMatrix(ctr, 2) & "' and Accessionnum = '" & .TextMatrix(ctr, 3) & "' and to_char(ResultDate,'MM/DD/YYYY HH24:MI') = '" & Format(.TextMatrix(ctr, 1), "mm/dd/yyyy HH:mm") & "' and EXAMNAME = '" & rstExams.Fields("InstExamName").Value & "'"
                                    myConnection.Execute Replace(temps, "<schema>", DBSchema)
                                End If
                            End If
                            rstExams.MoveNext
                        Loop
                        rstExams.MoveFirst
                    Else
                        MsgBox "Unable to transfer result. Click ok to continue", vbOKOnly + vbInformation, App.ProductName
                    End If
                Else
                    MsgBox "Unable to transfer result. Please check database settings. Click ok to continue", vbOKOnly + vbInformation, App.ProductName
                End If
            Next ctr
        End If
    Else
        MsgBox "No Exam to Transfer. Click ok to continue", vbOKOnly + vbInformation, App.ProductName
    End If
lblStatus.Caption = ""
rstResult.Requery
If rstResult.RecordCount = 0 Then
    frmSearchResult.Hide
Else
    Set msfgResult.DataSource = rstResult
    If temprow > .Rows - 1 Then .Row = .Rows - 1 Else .Row = temprow
    .ColSel = 3
    lblTitle.Caption = "Search Result - [" & msfgResult.Rows - 1 & " Record(s) Found" & "]"
    msfgResult.SetFocus
End If
End With
Screen.MousePointer = 0
End Sub

Private Sub OpenLocalConnection()
    Set eConnection = myConnection
End Sub

Private Function getConStr() As String
Dim tempstr As String
Dim mUserName As String
Dim mDatabase As String
Dim mPassword  As String
    tempstr = GetSetting("lab-i", "LOCAL", "UserName", "")
    If Len(Trim(tempstr)) > 0 Then
        mUserName = Blowfish.DecryptString(tempstr, myKey, True)
    Else
        mUserName = ""
    End If
    tempstr = GetSetting("lab-i", "LOCAL", "DATABASEPASSWORD", "")
    If Len(Trim(tempstr)) > 0 Then
        mPassword = Blowfish.DecryptString(tempstr, myKey, True)
    Else
        mPassword = ""
    End If
    tempstr = GetSetting("lab-i", "LOCAL", "DATABASE", "")
    If Len(Trim(tempstr)) > 0 Then
        mDatabase = Blowfish.DecryptString(tempstr, myKey, True)
    Else
        mDatabase = ""
    End If
    getConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mDatabase & ";User Id=" & mUserName & ";Password=;" & "Jet OLEDB:Database Password=" & mPassword & ";"
End Function

Public Property Let wTrans(ByVal vNewValue As String)
    yTrans = vNewValue
End Property

Private Function findRecorsetData(rstTemp As ADODB.Recordset, findString As String) As String
With rstTemp
    .MoveFirst
    .Find "[sequence #] = '" & findString & "'"
    If Not .EOF Then
        findRecorsetData = .Fields(1).Value
    Else
        findRecorsetData = ""
    End If
End With
End Function

Private Sub setIcons(CurrentRow As Long, Optional EndRow As Long = 0)
Dim ctr As Long
Dim mEndRow As Long
With msfgResult
.Redraw = False
If EndRow = 0 Then mEndRow = .Rows - 1 Else mEndRow = EndRow
For ctr = CurrentRow To mEndRow
    .Row = ctr
    .Col = 0
    .CellBackColor = Shape1.FillColor
    .CellPictureAlignment = 3
    If Len(Trim(.TextMatrix(ctr, 5))) > 0 Then
        If Format(.TextMatrix(ctr, 8), ">") = "TRUE" Then
            If Right(.TextMatrix(ctr, 5), 1) = "N" Then
                Set .CellPicture = imlImages.Overlay("new", 3)
            Else
                Set .CellPicture = imlImages.Overlay("printed", 2)
            End If
            .Col = 1
            .CellBackColor = Shape3.FillColor
            .Col = 2
            .CellBackColor = Shape3.FillColor
            .Col = 3
            .CellBackColor = Shape3.FillColor
        Else
            If Right(.TextMatrix(ctr, 5), 1) = "N" Then
                Set .CellPicture = imlImages.Overlay("new", 3)
            Else
                Set .CellPicture = imlImages.Overlay("saved", 1)
            End If
            .Col = 1
            .CellBackColor = Shape2.FillColor
            .Col = 2
            .CellBackColor = Shape2.FillColor
            .Col = 3
            .CellBackColor = Shape2.FillColor
        End If
    Else
'        Set .CellPicture = imlImages.Overlay(1, 2)
    End If
    If Len(Trim(.TextMatrix(ctr, 7))) > 0 Then
        .TextMatrix(ctr, 2) = .TextMatrix(ctr, 7)
    End If
    If Not .RowIsVisible(ctr) Then Exit For
Next ctr
.Row = CurrentRow
.Col = 1
.ColSel = 3
.Redraw = True
End With
End Sub

Private Sub PrintResults(Optional directprint As Boolean = True)
Dim ctrStart As Integer
Dim ctrLast As Integer
Dim ctr As Integer
With msfgResult
.Enabled = False
If .MouseRow > 0 Then
    If DBServerType = "LOCAL DATABASE" Then
        With msfgResult
            If .Row >= .RowSel Then
                ctrStart = .RowSel
                ctrLast = .Row
            Else
                ctrStart = .Row
                ctrLast = .RowSel
            End If
            ctr = ctrStart
            Do While ctr >= ctrStart And ctr <= ctrLast
                If Len(Trim(.TextMatrix(ctr, 5))) > 0 Then
                    If directprint Then
                        HostSA.ResultPrint Trim(.TextMatrix(ctr, 5))
                        lblStatus.Caption = "Printing: " & .TextMatrix(ctr, 2)
                    Else
                        HostSA.showResultPrintPreview Trim(.TextMatrix(ctr, 5))
                        lblStatus.Caption = "Previewing: " & .TextMatrix(ctr, 2)
                    End If
                Else
                    If Val(GetSetting("Lab-I", "Settings", "Autoprocess", "0")) = 1 Then
                        With msfgResult
                            HostSA.ResultAutoPrint GetSetting("Lab-I", "Settings", "ProfileID", "1"), InstrumentID, .TextMatrix(.Row, 3), .TextMatrix(.Row, 6), .TextMatrix(.Row, 1)
                        End With
                    End If
                End If
                ctr = ctr + 1
            Loop
        End With
    End If
    rstResult.Requery
    If rstResult.RecordCount = 0 Then
        frmSearchResult.Hide
    Else
        Set .DataSource = rstResult
        lblTitle.Caption = "Search Result - [" & msfgResult.Rows - 1 & " Record(s) Found" & "]"
        If .Enabled Then .SetFocus
        setIcons .TopRow ' ctrStart
        .RowSel = ctrLast
        .Redraw = False
        .Redraw = True
    End If
End If
.Enabled = True
End With
msfgResult_SelChange
End Sub

Private Sub MergeResults(myStr As String)
Dim ctr As Integer
Dim myResID As String
Dim arrResultID() As String
If InStr(1, myStr, "|") > 0 Then
    arrResultID = Split(myStr, "|")
    myResID = arrResultID(0)
    For ctr = 1 To UBound(arrResultID)
        If ctr > 1 Then
            myConnection.Execute "Update tbLabInstrumentRValues set ResultID = '" & myResID & "' where resultid = '" & arrResultID(ctr - 1) & "'"
            myConnection.Execute "Delete from tbLabInstrumentResults where resultid = '" & arrResultID(ctr - 1) & "'"
        End If
    Next ctr
End If
End Sub

Private Sub HighlightOver(ByRef obj As Object, Optional bgColor As Long = vbHighlight, _
    Optional ForeColor As Long = vbHighlightText, _
    Optional IsOn As Boolean = True)
    On Error GoTo ERR_HIGHLIGHT
    Dim Ctl As Control
    With obj
        If IsOn Then
            If (.Name <> "lblTitle" And .Name <> "lblStatus") Then
                .BackColor = bgColor
                .ForeColor = ForeColor
            End If
            For Each Ctl In Me
                With Ctl
                    If (TypeOf Ctl Is Label) Then
                        If .Name <> obj.Name Then
                            If (.Name <> "lblTitle" And .Name <> "lblStatus") Then
                                .BackColor = vbButtonFace
                                .ForeColor = vbNormal
                            End If
                        End If
                    End If
                End With
            Next Ctl
        Else
            For Each Ctl In Me
                With Ctl
                    .BackColor = vb3DFace
                    .ForeColor = vbNormal
                End With
            Next Ctl
        End If
    End With
EXIT_HIGHLIGHT:
    Exit Sub
ERR_HIGHLIGHT:
    If Err.Number = 438 Then
        Resume Next
    End If
End Sub

Private Function GetResultID() As String
Dim ctr, RowStart, RowEnd As Integer
Dim tempstr As String
tempstr = ""
With msfgResult
    If .Row <> .RowSel Then
        If .Row > .RowSel Then
            RowStart = .RowSel
            RowEnd = .Row
        ElseIf .Row < .RowSel Then
            RowStart = .Row
            RowEnd = .RowSel
        Else
            GetResultID = "FAIL"
            Exit Function
        End If
    End If
    For ctr = RowStart To RowEnd
        tempstr = tempstr & .TextMatrix(ctr, 10) & "|"
    Next ctr
    GetResultID = tempstr
End With
End Function

Private Function CanMerge() As Boolean
Dim ctr, RowStart, RowEnd As Integer
Dim pName, pSequence As String
CanMerge = True
With msfgResult
    If .Row <> .RowSel Then
        If .Row > .RowSel Then
            RowStart = .RowSel
            RowEnd = .Row
        ElseIf .Row < .RowSel Then
            RowStart = .Row
            RowEnd = .RowSel
        Else
            CanMerge = False
            Exit Function
        End If
    Else
            CanMerge = False
            Exit Function
    End If
    pName = .TextMatrix(RowStart, 2)
    pSequence = .TextMatrix(RowStart, 3)
    For ctr = RowStart To RowEnd
        If ctr > RowStart Then
            If Trim(.TextMatrix(ctr, 2)) <> Trim(pName) Then
                CanMerge = False
                Exit Function
            End If
            If Trim(.TextMatrix(ctr, 3)) <> Trim(pSequence) Then
                CanMerge = False
                Exit Function
            End If
        End If
    Next ctr
End With
End Function

Private Sub ResendVersa(mRefNum As String)
Dim rstTemp As ADODB.Recordset
Dim vsHeader As String
Dim vsDetails As String
Dim mVersa As Object
Dim ctr As Integer
vsHeader = ""
vsDetails = ""
If Val(GetSetting("Lab-I", "Versa", "EnablePlugin", "0")) = 1 Then
    If MsgBox("Are you sure you want to resend the results to HIS?", vbQuestion + vbYesNo + vbDefaultButton2, "Resend Result") = vbYes Then
        Set mVersa = CreateObject(Replace(Format(GetSetting("Lab-I", "VerSa", "PlugIN", ""), ">"), ".DLL", ".mFunctions"))
        Set rstTemp = myConnection.Execute("vs_ValuesSelect '" & mRefNum & "'")
        If rstTemp.RecordCount > 0 Then
            With rstTemp
                .MoveFirst
                vsHeader = .Fields("Patient.RefNum").Value & "|" & .Fields("LastName").Value & "|" & .Fields("FirstName").Value & "|" & .Fields("MiddleName").Value & _
                "|" & .Fields("Age").Value & "|" & .Fields("Sex").Value & "|" & .Fields("BirthDate").Value & "|" & .Fields("Ward").Value & "|" & .Fields("Room").Value & _
                "|" & .Fields("ResultDate").Value & "|" & .Fields("RequestDate").Value & "|" & .Fields("Physician").Value & "|" & .Fields("MedTech").Value & _
                "|" & .Fields("ProfileExams.InstrumentID").Value & "|" & .Fields("Remarks").Value & "|" & .Fields("ProfileName").Value & "|" & .Fields("ClassName").Value & _
                "|" & .Fields("Pathologist").Value & "|" & .Fields("PatientID").Value & "|" & .Fields("SpecimenID").Value & "|" & .Fields("AdditionalProfiles").Value & _
                "|" & .Fields("CivilStatus").Value & "|" & .Fields("Specimen").Value
                ctr = 1
                Do While Not .EOF
                    vsDetails = vsDetails & .Fields("Patient.RefNum").Value & "|" & .Fields("ProfileExams.ExamID").Value & "|" & .Fields("InstrumentResult").Value & "|" & .Fields("Result").Value & _
                    "|" & .Fields("NValuesDisplay").Value & "|" & .Fields("UNIT").Value & "|" & .Fields("Status").Value & "|" & .Fields("Recheck").Value & _
                    "|" & Trim(.Fields("DisplayName").Value) & "|" & .Fields("ProfileExams.Sortnum").Value & "|" & .Fields("NValuesSIDisplay").Value & "|" & .Fields("UnitSI").Value & _
                    "|" & .Fields("ResultSI").Value & "|" & .Fields("Profile").Value & "|" & .Fields("ExamNotes").Value & "|" & .Fields("InstExamName").Value & "<SP>"
                    ctr = ctr + 1
                    .MoveNext
                Loop
                If Len(vsHeader) > 0 And Len(vsDetails) > 0 Then
                    mVersa.ProcessVSP vsHeader, vsDetails
                    MsgBox "Upload Done", vbInformation + vbOKOnly, "Resend Result"
                End If
            End With
        End If
    End If
End If
End Sub

Private Sub tbUpload_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "send"
        ResendVersa msfgResult.TextMatrix(msfgResult.Row, 5)
End Select
End Sub
