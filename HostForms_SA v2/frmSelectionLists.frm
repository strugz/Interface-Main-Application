VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSelectionLists 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selection Lists"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   9570
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
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
      Height          =   495
      Left            =   10200
      TabIndex        =   0
      Top             =   1650
      Width           =   2775
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Build File - [Selection Lists]"
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
         TabIndex        =   1
         Top             =   150
         Width           =   2595
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
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
      Height          =   6075
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   9495
      Begin VB.TextBox txtTag 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6630
         TabIndex        =   18
         Top             =   5130
         Width           =   2145
      End
      Begin VB.CommandButton cmdDeleteCat 
         Caption         =   "Delete Categories"
         Height          =   405
         Left            =   1770
         TabIndex        =   15
         Top             =   5550
         Width           =   1575
      End
      Begin VB.CommandButton cmdDown 
         Height          =   975
         Left            =   8940
         Picture         =   "frmSelectionLists.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2280
         Width           =   435
      End
      Begin VB.CommandButton cmdUp 
         Height          =   975
         Left            =   8940
         Picture         =   "frmSelectionLists.frx":0A02
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1260
         Width           =   435
      End
      Begin VB.CommandButton cmdAddCategories 
         Caption         =   "Add Categories"
         Height          =   405
         Left            =   150
         TabIndex        =   11
         Top             =   5550
         Width           =   1605
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   405
         Left            =   7560
         TabIndex        =   10
         Top             =   5520
         Width           =   1275
      End
      Begin VB.ListBox lstType 
         Appearance      =   0  'Flat
         Height          =   5070
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3255
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Default         =   -1  'True
         Height          =   435
         Left            =   6240
         TabIndex        =   6
         Top             =   5520
         Width           =   1275
      End
      Begin VB.TextBox txtValues 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3810
         TabIndex        =   5
         Top             =   5100
         Width           =   2145
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   4695
         Left            =   3420
         TabIndex        =   9
         Top             =   360
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   8281
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   12615680
         ForeColorFixed  =   16777215
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
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Value:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   6060
         TabIndex        =   17
         Top             =   5160
         Width           =   465
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Item:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   3480
         TabIndex        =   16
         Top             =   5130
         Width           =   330
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
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   1
         Left            =   3390
         TabIndex        =   12
         Top             =   2730
         Width           =   5385
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Categories:"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   150
         Width           =   825
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
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
      Height          =   705
      Index           =   1
      Left            =   30
      TabIndex        =   3
      Top             =   6150
      Width           =   9495
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   375
         Left            =   8010
         TabIndex        =   4
         Top             =   210
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmSelectionLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstType As ADODB.Recordset
Private rstValues As ADODB.Recordset
Private yClassForms As New ClassForms

Private Sub cmdAdd_Click()
    If Len(txtValues.Text) > 0 Then
        myConnection.Execute "Insert into tbSelectionValues (selType, selValues, selTag) Select '" & Replace(lstType.Text, "'", "''") & "','" & Replace(txtValues.Text, "'", "''") & "','" & Replace(txtTag.Text, "'", "''") & "'"
    End If
    findValues lstType.Text
    txtValues.Text = ""
    If txtValues.Visible Then txtValues.SetFocus
End Sub

Private Sub cmdAddCategories_Click()
Dim tempStr As String
    tempStr = yClassForms.showInputBox("Enter Category", , , , , True)
    If Len(Trim(tempStr)) > 0 Then
        myConnection.Execute "Insert into tbSelectionType (seltype) select '" & Replace(tempStr, "'", "''") & "'"
        Set rstType = myConnection.Execute("select * from tbSelectionType where status = '1'")
        filllstType
        msfgResult.ColWidth(0) = msfgResult.Width - (100 * 3)
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
myConnection.Execute "Delete from tbSelectionValues where seltype = '" & Replace(lstType.Text, "'", "''") & "' and selvalues = '" & Replace(msfgResult.TextMatrix(msfgResult.Row, 0), "'", "''") & "';"
findValues lstType.Text
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdDeleteCat_Click()
myConnection.Execute "Delete from tbSelectionType where seltype = '" & Replace(lstType.Text, "'", "''") & "'"
rstType.Requery
filllstType
End Sub

Private Sub Form_Load()
    setFormColors Me
    On Error Resume Next
    Set rstType = myConnection.Execute("select * from tbSelectionType where status = '1'")
    filllstType
    If rstType.RecordCount > 0 Then
        msfgResult.ColWidth(0) = (msfgResult.Width * 0.6) - (60 * 3)
        msfgResult.ColWidth(1) = (msfgResult.Width * 0.2) - (20 * 3)
        msfgResult.ColWidth(2) = (msfgResult.Width * 0.2) - (20 * 3)
        msfgResult.TextMatrix(0, 2) = "Order"
    End If
End Sub

Private Sub filllstType()
Dim ctr As Integer
With lstType
    .Clear
    If rstType.RecordCount > 0 Then
        rstType.MoveFirst
        Do While Not rstType.EOF
            .AddItem rstType.Fields("SelType").Value
            rstType.MoveNext
        Loop
        .ListIndex = 0
    End If
End With
End Sub

Private Sub findValues(findString As String)
If DBServerType = "LOCAL DATABASE" Then
    Set rstValues = myConnection.Execute("Select selValues as [Values], selTag as [Tag], selsortorder from tbSelectionValues where seltype = '" & Replace(findString, "'", "''") & "'")
ElseIf DBServerType = "MYSQL SERVER" Then
    Set rstValues = myConnection.Execute("Select selValues ""Values"", selTag ""Tag"", selsortorder from tbSelectionValues where seltype = '" & Replace(findString, "'", "''") & "'")
ElseIf DBServerType = "MSSQL SERVER" Then
    Set rstValues = myConnection.Execute("Select selValues ""Values"", selTag ""Tag"", selsortorder from tbSelectionValues where seltype = '" & Replace(findString, "'", "''") & "'")
End If
    If rstValues.RecordCount = 0 Then
        msfgResult.Visible = False
        cmdDelete.Enabled = False
        cmdUp.Visible = False
        cmdDown.Visible = False
    Else
        msfgResult.Visible = True
        cmdDelete.Enabled = True
        cmdUp.Visible = True
        cmdDown.Visible = True
        Set msfgResult.DataSource = rstValues
    End If
  
End Sub

Private Sub lstType_Click()
    findValues lstType.Text
End Sub



Private Sub txtValues_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
