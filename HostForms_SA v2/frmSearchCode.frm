VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSearchCode 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search Code"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   8280
      TabIndex        =   4
      Top             =   2520
      Width           =   7485
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Code"
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
         TabIndex        =   5
         Top             =   150
         Width           =   1215
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   2
      Left            =   30
      TabIndex        =   8
      Top             =   30
      Width           =   7485
      Begin VB.TextBox txtFilter 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   540
         TabIndex        =   0
         Top             =   180
         Width           =   6825
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter"
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
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   210
         Width           =   345
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   3705
      Index           =   0
      Left            =   30
      TabIndex        =   6
      Top             =   600
      Width           =   7485
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   3405
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   6006
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   12615680
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
      End
      Begin VB.Label lblText 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No Record Found..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   1
         Left            =   60
         TabIndex        =   10
         Top             =   1410
         Width           =   7395
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   705
      Index           =   1
      Left            =   30
      TabIndex        =   7
      Top             =   4350
      Width           =   7485
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
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
         Left            =   5970
         TabIndex        =   3
         Top             =   210
         Width           =   1395
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         Default         =   -1  'True
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
         Left            =   4440
         TabIndex        =   2
         Top             =   210
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmSearchCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ySearchValue As String
Private rstData As ADODB.Recordset
Private yshowFilter As Boolean

Private Sub cmdClose_Click()
    ReturnValueStr = ""
    ReturnValueStr2 = ""
    Unload Me
End Sub

Public Property Let wSearchValue(ByVal vNewValue As String)
    ySearchValue = vNewValue
End Property

Public Property Let wRecordSet(ByVal vNewValue As ADODB.Recordset)
    Set rstData = vNewValue
End Property

Private Sub cmdSelect_Click()
With msfgResult
    ReturnValueStr = .TextMatrix(.Row, 1)
    ReturnValueStr2 = .TextMatrix(.Row, 2)
End With
    Unload Me
End Sub

Private Sub Form_Load()
setFormColors Me
Set msfgResult.DataSource = rstData
setFlexGrid
If yshowFilter Then
    msfgResult.Top = 180
    msfgResult.Height = 3405
    fraBody(0).Top = 720
    fraBody(0).Height = 3705
    fraBody(2).Visible = True
Else
    msfgResult.Top = 180
    msfgResult.Height = 3765
    fraBody(0).Top = 270
    fraBody(0).Height = 4155
    fraBody(2).Visible = False
End If
End Sub

Private Sub setFlexGrid()
With msfgResult
'    .Height = .CellHeight * 16
    .ColWidth(0) = .Width * 0.05 - (30 * 3)
    .ColWidth(1) = .Width * 0.3 - (30 * 3)
    .ColWidth(2) = .Width * 0.65 - (40 * 3)
    .ColSel = 2
End With
End Sub

Private Sub lblTitle_Change()
    Me.Caption = lblTitle.Caption
End Sub

Private Sub msfgResult_DblClick()
    cmdSelect_Click
End Sub

Public Property Let wShowFilter(ByVal vNewValue As Boolean)
    yshowFilter = vNewValue
End Property

Private Sub txtFilter_Change()
Dim tempStr As String
    tempStr = "[" & rstData.Fields(1).Name & "] like '%" & txtFilter.Text & "%'"
    If Len(Trim(txtFilter)) > 0 Then
        rstData.Filter = tempStr
    Else
        rstData.Filter = ""
    End If
    If rstData.RecordCount = 0 Then
        Set msfgResult.Recordset = Nothing
        msfgResult.Visible = False
        cmdSelect.Enabled = False
    Else
        Set msfgResult.Recordset = rstData
        msfgResult.Visible = True
        cmdSelect.Enabled = True
    End If
End Sub

Private Sub txtFilter_GotFocus()
    highLight
End Sub

Private Sub txtFilter_KeyDown(KeyCode As Integer, Shift As Integer)
With msfgResult
    Select Case KeyCode
        Case vbKeyDown
            If .Visible Then
                If .Row + 1 <> .Rows Then
                    .Row = .Row + 1
                    .ColSel = 2
                End If
            End If
        Case vbKeyUp
            If .Visible Then
                If .Row <> 1 Then
                    .Row = .Row - 1
                    .ColSel = 2
                End If
            End If
    End Select
End With
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
