VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSpecimen 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Specimen"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6090
      TabIndex        =   0
      Top             =   2310
      Width           =   5115
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Build File - [Specimen]"
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
         Width           =   2115
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   5205
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   30
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
         Left            =   960
         TabIndex        =   4
         Top             =   4680
         Width           =   3375
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
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Specimen:"
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
         Left            =   120
         TabIndex        =   6
         Top             =   4740
         Width           =   750
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   30
      TabIndex        =   7
      Top             =   5250
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
Attribute VB_Name = "frmSpecimen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstMedtech As ADODB.Recordset

Private Sub cmdAdd_Click()
Dim rstTemp As ADODB.Recordset
Dim tempStr As String
tempStr = Trim(Replace(txtMedtech.Text, "'", "''"))
Set rstTemp = myConnection.Execute("Select * from tblabpathologists where pathologist = '" & tempStr & "'")
If rstTemp.RecordCount > 0 Then
    MsgBox "Name already exists.", vbOKOnly + vbInformation, "Pathologist"
Else
    If Len(Trim(txtMedtech)) > 0 Then
        myConnection.Execute "Insert into tblabSpecimen (Specimen) select '" & Replace(txtMedtech.Text, "'", "''") & "',True"
        MsgBox "Record has been successfully saved.", vbOKOnly + vbInformation, "Med.Tech."
    End If
    txtMedtech = ""
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
Set rstMedtech = myConnection.Execute("Select Specimen from tblabSpecimen")
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


Private Sub msfgResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete And Shift = 1 Then
    If MsgBox("Are you sure you want to remove this item?", vbYesNo + vbQuestion + vbDefaultButton2, "Delete Item") = vbYes Then
        myConnection.Execute "Delete from tblabSpecimen where Specimen = '" & Replace(msfgResult.TextMatrix(msfgResult.Row, msfgResult.Col), "'", "''") & "'"
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

Private Sub txtmedtech_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub



