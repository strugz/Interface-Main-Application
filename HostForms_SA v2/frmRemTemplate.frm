VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRemTemplate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reamarks Keywords"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5220
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   5220
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6540
      TabIndex        =   0
      Top             =   2790
      Width           =   5115
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Build File - [Remarks Keywords]"
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
         Width           =   3060
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   5205
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   5115
      Begin VB.TextBox txtmedtech 
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
         Height          =   315
         Left            =   870
         TabIndex        =   4
         Top             =   4680
         Width           =   3465
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
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
         Left            =   4350
         TabIndex        =   3
         Top             =   4710
         Width           =   675
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   4395
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   7752
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
         Caption         =   "Keyword"
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
         TabIndex        =   6
         Top             =   4740
         Width           =   675
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   60
      TabIndex        =   7
      Top             =   5310
      Width           =   5115
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
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
         Left            =   3900
         TabIndex        =   8
         Top             =   180
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmRemTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstMedtech As ADODB.Recordset

Private Sub cmdAdd_Click()
Dim rstTemp As ADODB.Recordset
Dim tempStr As String
tempStr = Trim(Replace(txtmedtech.Text, "'", "''"))
Set rstTemp = myConnection.Execute("Select * from tbLabRemTemplate where templates = '" & tempStr & "'")
If rstTemp.RecordCount > 0 Then
    MsgBox "Name already exists.", vbOKOnly + vbInformation, "Med.Tech"
Else
    If Len(Trim(txtmedtech.Text)) > 0 Then
        myConnection.Execute "Insert into tbLabRemTemplate (templates) select '" & Replace(txtmedtech.Text, "'", "''") & "'"
        MsgBox "Record has been successfully saved.", vbOKOnly + vbInformation, "Med.Tech."
    End If
    txtmedtech = ""
    rstMedtech.Requery
    If rstMedtech.RecordCount > 0 Then
        msfgResult.Visible = True
        Set msfgResult.DataSource = rstMedtech
        setFlexGrid
    Else
        msfgResult.Visible = False
    End If
End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
setFormColors Me
If DBServerType = "LOCAL DATABASE" Then
     Set rstMedtech = myConnection.Execute("Select templates as [Keywords] from tbLabRemTemplate")
ElseIf DBServerType = "MYSQL SERVER" Then
    Set rstMedtech = myConnection.Execute("Select templates ""Keywords"" from tbLabRemTemplate")
ElseIf DBServerType = "MSSQL SERVER" Then
    Set rstMedtech = myConnection.Execute("Select templates ""Keywords"" from tbLabRemTemplate")
End If

If rstMedtech.RecordCount > 0 Then
    msfgResult.Visible = True
    Set msfgResult.DataSource = rstMedtech
    setFlexGrid
Else
    msfgResult.Visible = False
End If
End Sub

Private Sub setFlexGrid()
With msfgResult
    .Height = .CellHeight * 18
    .ColWidth(0) = .Width * 1 - (100 * 3)
End With
End Sub

Private Sub lblTitle_Change()
    Me.Caption = lblTitle.Caption
End Sub

Private Sub msfgResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete And Shift = 1 Then
    If MsgBox("Are you sure you want to remove this item?", vbYesNo + vbQuestion + vbDefaultButton2, "Delete Item") = vbYes Then
        myConnection.Execute "Delete from tbLabRemTemplate where templates = '" & Replace(msfgResult.TextMatrix(msfgResult.Row, msfgResult.Col), "'", "''") & "'"
        rstMedtech.Requery
        If rstMedtech.RecordCount > 0 Then
            msfgResult.Visible = True
            Set msfgResult.DataSource = rstMedtech
            setFlexGrid
        Else
            msfgResult.Visible = False
        End If
    End If
End If
End Sub

