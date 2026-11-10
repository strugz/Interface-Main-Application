VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmMedtech 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Med. Tech."
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   9210
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   11160
      TabIndex        =   0
      Top             =   2790
      Width           =   5115
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Build File - [Medical Technologists]"
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
         Width           =   3300
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   5895
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   9135
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
         Height          =   975
         Left            =   6810
         TabIndex        =   13
         Top             =   4710
         Width           =   1185
      End
      Begin VB.TextBox txtLicExpire 
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
         Left            =   2070
         TabIndex        =   10
         Top             =   5370
         Width           =   3465
      End
      Begin VB.TextBox txtLicense 
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
         Left            =   2070
         TabIndex        =   9
         Top             =   5040
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
         Height          =   975
         Left            =   5580
         TabIndex        =   8
         Top             =   4710
         Width           =   1185
      End
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
         Left            =   2070
         MaxLength       =   100
         TabIndex        =   6
         Top             =   4710
         Width           =   3465
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   4395
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   8955
         _ExtentX        =   15796
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
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NO RECORD FOUND."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   2070
         Width           =   8865
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Expiration:"
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
         Index           =   2
         Left            =   1260
         TabIndex        =   12
         Top             =   5370
         Width           =   750
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "License #:"
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
         Index           =   1
         Left            =   1260
         TabIndex        =   11
         Top             =   5070
         Width           =   750
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Med.Tech:"
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
         Left            =   1260
         TabIndex        =   7
         Top             =   4770
         Width           =   750
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   1
      Left            =   30
      TabIndex        =   3
      Top             =   5820
      Width           =   9135
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
         Left            =   3540
         TabIndex        =   4
         Top             =   180
         Width           =   2025
      End
   End
End
Attribute VB_Name = "frmMedtech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstMedtech As ADODB.Recordset

Private Sub cmdAdd_Click()
Dim rstTemp As ADODB.Recordset
Dim tempStr As String
Dim strLicense As String
Dim strLicExpire As String
Dim strQry As String
With msfgResult
strQry = ""
tempStr = Trim(Replace(txtMedtech.Text, "'", "''"))
strLicense = Trim(Replace(txtLicense.Text, "'", "''"))
strLicExpire = Trim(Replace(txtLicExpire.Text, "'", "''"))
If cmdAdd.Caption = "Add" Then
    Set rstTemp = myConnection.Execute("Select * from tbMedtech where mtname = '" & tempStr & "' order by mtname")
    If rstTemp.RecordCount > 0 Then
        MsgBox "Name already exists.", vbOKOnly + vbInformation, "Med.Tech"
    Else
        If Len(Trim(txtMedtech.Text)) > 0 Then
            myConnection.Execute "Insert into tbmedtech (mtname, mtLicense, mtLicExpire) select '" & tempStr & "','" & strLicense & "','" & strLicExpire & "'"
            MsgBox "Record has been successfully saved.", vbOKOnly + vbInformation, "Med.Tech."
        End If
    End If
Else
    If Trim(.TextMatrix(.Row, 0)) <> Trim(txtMedtech.Text) Then
        Set rstTemp = myConnection.Execute("Select * from tbMedtech where mtname = '" & Replace(txtMedtech.Text, "'", "''") & "' order by mtname")
        If rstTemp.RecordCount > 0 Then
            MsgBox "Name already exists.", vbOKOnly + vbInformation, "Med.Tech"
        Else
            strQry = "Update tbMedTech set mtName='" & tempStr & "', mtLicense='" & strLicense & "', mtLicExpire='" & strLicExpire & "' where mtName = '" & Trim(.TextMatrix(.Row, 0)) & "'"
        End If
    Else
        strQry = "Update tbMedTech set mtName='" & tempStr & "', mtLicense='" & strLicense & "', mtLicExpire='" & strLicExpire & "' where mtName = '" & Trim(.TextMatrix(.Row, 0)) & "'"
    End If
    If Len(Trim(strQry)) > 0 Then myConnection.Execute strQry
    cmdAdd.Caption = "Add"
    cmdDelete.Caption = "Delete"
    cmdClose.Cancel = True
    .Enabled = True
End If
txtMedtech.Text = ""
txtLicense.Text = ""
txtLicExpire.Text = ""
rstMedtech.Requery
If rstMedtech.RecordCount > 0 Then
    msfgResult.Visible = True
    Set msfgResult.DataSource = rstMedtech
    setFlexGrid
Else
    msfgResult.Visible = False
    cmdDelete.Enabled = False
End If
End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdDelete_Click()
Dim i As Integer
If cmdDelete.Caption = "Cancel" Then
    txtMedtech.Text = ""
    txtLicense.Text = ""
    txtLicExpire.Text = ""
    cmdDelete.Caption = "Delete"
    cmdClose.Cancel = True
    cmdAdd.Caption = "Add"
    msfgResult.Enabled = True
Else
    If MsgBox("Are you sure you want to remove this item(s)?", vbYesNo + vbQuestion + vbDefaultButton2, "Delete Item") = vbYes Then
        With msfgResult
            For i = .Row To .RowSel
                myConnection.Execute "Delete from tbmedtech where mtName = '" & Replace(.TextMatrix(i, 0), "'", "''") & "'"
            Next i
            rstMedtech.Requery
            If rstMedtech.RecordCount > 0 Then
                .Visible = True
                Set .DataSource = rstMedtech
                setFlexGrid
            Else
                .Visible = False
                cmdDelete.Enabled = False
            End If
        End With
    End If
End If
End Sub

Private Sub Form_Load()
Dim rstTemp As ADODB.Recordset
Dim ctr As Integer
Dim hasLicense As Boolean
Dim hasLicExpire As Boolean
setFormColors Me
hasLicense = False
hasLicExpire = False
Set rstTemp = myConnection.Execute("Select * From tbMedTech")
With rstTemp
    For ctr = 0 To rstTemp.Fields.Count - 1
        If rstTemp.Fields(ctr).Name = "mtLicense" Then hasLicense = True
        If rstTemp.Fields(ctr).Name = "mtLicExpire" Then hasLicExpire = True
    Next ctr
End With
If Not hasLicense Then myConnection.Execute "ALTER TABLE tbMedTech ADD COLUMN mtLicense TEXT(25)"
If Not hasLicExpire Then myConnection.Execute "ALTER TABLE tbMedTech ADD COLUMN mtLicExpire TEXT(25)"

If SectionCode = "PULMO" Then
    lblTitle.Caption = "Build File - [Respiratory Therapist]"
    lblText(0).Caption = "Res. Ther.:"
    If DBServerType = "LOCAL DATABASE" Then
        Set rstMedtech = myConnection.Execute("Select mtname, mtLicense, mtLicExpire as [Respiratory Therapist] from tbMedtech")
    ElseIf DBServerType = "MYSQL SERVER" Then
        Set rstMedtech = myConnection.Execute("Select mtname, mtLicense, mtLicExpire as ""Respiratory Therapist"" from tbMedtech")
    End If
Else
    lblTitle.Caption = "Build File - [Medical Technologists]"
    lblText(0).Caption = "Med. Tech.:"
    If DBServerType = "LOCAL DATABASE" Then
        Set rstMedtech = myConnection.Execute("Select mtname, mtLicense, mtLicExpire as [Medical Technologist] from tbMedtech")
    ElseIf DBServerType = "MYSQL SERVER" Then
        Set rstMedtech = myConnection.Execute("Select mtname, mtLicense, mtLicExpire as ""Medical Technologist"" from tbMedtech")
    ElseIf DBServerType = "MSSQL SERVER" Then
        Set rstMedtech = myConnection.Execute("Select mtname, mtLicense, mtLicExpire as ""Medical Technologist"" from tbMedtech")
    End If
End If
If rstMedtech.RecordCount > 0 Then
    msfgResult.Visible = True
    Set msfgResult.DataSource = rstMedtech
    setFlexGrid
Else
    msfgResult.Visible = False
    cmdDelete.Enabled = False
End If
End Sub

Private Sub setFlexGrid()
With msfgResult
    .Height = .CellHeight * 18
    .ColWidth(0) = .Width * 0.6 - 100
    .ColWidth(1) = .Width * 0.2 - 100
    .ColWidth(2) = .Width * 0.2 - 100
End With
End Sub

Private Sub lblTitle_Change()
    Me.Caption = lblTitle.Caption
End Sub

Private Sub msfgResult_DblClick()
With msfgResult
    If Len(Trim(.TextMatrix(.Row, 0))) > 0 Then
        .Enabled = False
        txtMedtech.Text = .TextMatrix(.Row, 0)
        txtLicense.Text = .TextMatrix(.Row, 1)
        txtLicExpire.Text = .TextMatrix(.Row, 2)
        cmdAdd.Caption = "Update"
        cmdDelete.Caption = "Cancel"
        cmdDelete.Cancel = True
    End If
End With
End Sub

Private Sub msfgResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete And Shift = 1 Then
    If MsgBox("Are you sure you want to remove this item?", vbYesNo + vbQuestion + vbDefaultButton2, "Delete Item") = vbYes Then
        myConnection.Execute "Delete from tbmedtech where mtName = '" & Replace(msfgResult.TextMatrix(msfgResult.Row, msfgResult.Col), "'", "''") & "'"
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
