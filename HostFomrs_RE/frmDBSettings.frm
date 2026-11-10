VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmReqRates 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rates"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   10425
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDiscount 
      BackColor       =   &H00A66B2B&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   15
      TabIndex        =   10
      Top             =   -120
      Width           =   10395
      Begin VB.CommandButton cmdCancelDiscount 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2145
         TabIndex        =   3
         Top             =   1500
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmdSaveDiscount 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3120
         TabIndex        =   2
         Top             =   1500
         Width           =   915
      End
      Begin VB.TextBox txtDiscountName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1545
         MaxLength       =   50
         TabIndex        =   1
         Top             =   975
         Width           =   2475
      End
      Begin VB.TextBox txtDiscountID 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   390
         Left            =   1545
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   495
         Width           =   2475
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgReqDiscounts 
         Height          =   2070
         Left            =   4380
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   270
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   3651
         _Version        =   393216
         FixedCols       =   0
         GridColorFixed  =   0
         FocusRect       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount ID:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   1
         Left            =   375
         TabIndex        =   13
         Top             =   570
         Width           =   1320
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Type:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   2
         Left            =   375
         TabIndex        =   12
         Top             =   1035
         Width           =   1320
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Discounts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   11
         Top             =   225
         Width           =   1320
      End
   End
   Begin VB.Frame fraRates 
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
      Height          =   5655
      Left            =   15
      TabIndex        =   9
      Top             =   2295
      Width           =   10395
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8910
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   4980
         Width           =   1350
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7485
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   4980
         Width           =   1350
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1500
         Width           =   1920
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1050
         Width           =   1920
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1350
         TabIndex        =   4
         Top             =   600
         Width           =   1920
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4380
         MaxLength       =   50
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   165
         Width           =   5880
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   4275
         Left            =   3375
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   600
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   7541
         _Version        =   393216
         FixedCols       =   0
         GridColorFixed  =   0
         FocusRect       =   0
         HighLight       =   2
         GridLinesFixed  =   1
         ScrollBars      =   2
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   6
         Left            =   375
         TabIndex        =   20
         Top             =   1530
         Width           =   1320
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Name:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   5
         Left            =   375
         TabIndex        =   19
         Top             =   1065
         Width           =   1320
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate ID:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   4
         Left            =   375
         TabIndex        =   18
         Top             =   615
         Width           =   1320
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Type:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   3
         Left            =   3390
         TabIndex        =   15
         Top             =   210
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmReqRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intRowCtr As Integer
Private intColCtr As Integer
Private errfound
Private saveFields As String

Private rsDiscounts As New ADODB.Recordset
Private rsRates As New ADODB.Recordset

Private Sub cmdCancelDiscount_Click()
    msfgReqDiscounts.Enabled = True
    fraRates.Enabled = True
    txtDiscountID.Text = ""
    txtDiscountName.Text = ""
    cmdSaveDiscount.Caption = "Save"
    cmdCancelDiscount.Visible = False
    cmdCancelDiscount.Enabled = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSaveDiscount_Click()
    Dim rsInsertDiscounts As New ADODB.Recordset
    
    With msfgReqDiscounts
        If cmdSaveDiscount.Caption = "Save" Then
            If .TextMatrix(1, 0) = "D" Then
                strSQLStatement = "EXEC sp_InsertExamDiscount '" & .TextMatrix(1, 1) & "'"
                Set rsInsertDiscounts = myConnection.Execute(strSQLStatement)
                setmsfgReqDiscounts
            End If
        
            If Len(Trim(txtDiscountName.Text)) > 0 Then
                strSQLStatement = "EXEC sp_InsertExamDiscount '" & Trim(txtDiscountName.Text) & "'"
                Set rsInsertDiscounts = myConnection.Execute(strSQLStatement)
                setmsfgReqDiscounts
            End If
        Else
            strSQLStatement = "EXEC sp_UpdateExamDiscount " & txtDiscountID.Text & ",'" & txtDiscountName.Text & "'"
            Set rsInsertDiscounts = myConnection.Execute(strSQLStatement)
            setmsfgReqDiscounts
            .Enabled = True
            cmdSaveDiscount.Caption = "Save"
            cmdCancelDiscount.Visible = False
            cmdCancelDiscount.Enabled = False
        End If
        txtDiscountID.Text = ""
        txtDiscountName.Text = ""
    End With
    
    fraRates.Enabled = True

End Sub

Private Sub Form_Load()
    setFormColors Me
    fraDiscount.BackColor = GetSetting("LMS", "FormColors", "BodyFrame", &H6B4B01)
    fraDiscount.ForeColor = GetSetting("LMS", "FormColors", "TitleText", &HFFFFFF)

    fraRates.BackColor = GetSetting("LMS", "FormColors", "BodyFrame", &H6B4B01)
    fraRates.ForeColor = GetSetting("LMS", "FormColors", "TitleText", &HFFFFFF)

    msfgReqDiscounts.BackColorFixed = GetSetting("LMS", "FormColors", "TitleBar", &HA66B2B)
    msfgReqDiscounts.ForeColorFixed = GetSetting("LMS", "FormColors", "TitleText", &HFFFFFF)
    
    setmsfgReqDiscounts
    setmsfgRates
End Sub

Private Sub msfgReqDiscounts_DblClick()
    With msfgReqDiscounts
        If .Row > 1 Then
            fraRates.Enabled = False
            txtDiscountID.Text = .TextMatrix(.Row, 0)
            txtDiscountName.Text = .TextMatrix(.Row, 1)
            cmdSaveDiscount.Caption = "Update"
            cmdCancelDiscount.Visible = True
            cmdCancelDiscount.Enabled = True
            .Enabled = False
        End If
        txtDiscountName.SetFocus
    End With
End Sub

Private Sub txtDiscountName_GotFocus()
    highLight
End Sub

Private Sub txtDiscountName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub setmsfgReqDiscounts()
   
    With msfgReqDiscounts
        strSQLStatement = "EXEC sp_GetDiscounts_Select"
        Set rsDiscounts = myConnection.Execute(strSQLStatement)
        
        Set .DataSource = rsDiscounts
        
        .TextArray(0) = "ID"
        .TextArray(1) = "Discount Type"
            
        .ColWidth(0) = 500
        .ColWidth(1) = 5000
            
        If rsDiscounts.RecordCount = 0 Then
            .Rows = 2
            .Row = 1
            .TextMatrix(.Row, 0) = "D"
            .TextMatrix(.Row, 1) = "NONE"
        End If
        
    End With
End Sub

Private Sub setmsfgRates()
    
    With msfgResult
        strSQLStatement = "EXEC sp_GetRates_Select '" & InstrumentID & "'"
        Set rsRates = myConnection.Execute(strSQLStatement)
        
        If rsRates.RecordCount > 0 Then
            rsRates.Filter = "DiscountID = " & msfgReqDiscounts.TextMatrix(msfgReqDiscounts.Row, 0)
        End If
        
        Set .DataSource = rsRates
        
        .TextArray(1) = "ID"
        .TextArray(4) = "Exam Name"
        
        .ColWidth(0) = 0
        .ColWidth(1) = 500
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColWidth(4) = 4000
        .ColWidth(5) = 2000
    End With
End Sub


