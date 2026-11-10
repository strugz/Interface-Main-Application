VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmNormalValues 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Normal Values Items"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11730
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11730
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   12720
      TabIndex        =   11
      Top             =   1410
      Width           =   3345
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Build File - [Normal Values]"
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
         TabIndex        =   12
         Top             =   150
         Width           =   2565
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   1185
      Index           =   0
      Left            =   30
      TabIndex        =   13
      Top             =   60
      Width           =   3465
      Begin VB.TextBox txtClassID 
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
         Left            =   1260
         MaxLength       =   8
         TabIndex        =   0
         Top             =   660
         Width           =   555
      End
      Begin VB.TextBox txtInstrumentID 
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
         Left            =   1260
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   270
         Width           =   2115
      End
      Begin VB.Label txtClassName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   1830
         TabIndex        =   25
         Top             =   660
         Width           =   1545
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Classification:"
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
         Left            =   240
         TabIndex        =   19
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument ID:"
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
         Left            =   270
         TabIndex        =   18
         Top             =   330
         Width           =   975
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6045
      Index           =   1
      Left            =   30
      TabIndex        =   14
      Top             =   1290
      Width           =   3465
      Begin RichTextLib.RichTextBox txtDisplayName 
         Height          =   975
         Left            =   1230
         TabIndex        =   37
         Top             =   1830
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   1720
         _Version        =   393217
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmNormalvalues.frx":0000
      End
      Begin VB.CommandButton cmdDeleteItems 
         Caption         =   "Delete Items..."
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   4860
         Width           =   3225
      End
      Begin VB.CommandButton cmdCopyFrom 
         Caption         =   "Copy normal values from class..."
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   4440
         Width           =   3225
      End
      Begin VB.TextBox txtHighRangeSI 
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
         Left            =   1500
         TabIndex        =   6
         Top             =   3210
         Width           =   1845
      End
      Begin VB.TextBox txtLowRangeSI 
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
         Left            =   1500
         TabIndex        =   5
         Top             =   2820
         Width           =   1845
      End
      Begin VB.TextBox txtUnitSI 
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
         Left            =   1500
         TabIndex        =   7
         Top             =   3600
         Width           =   1845
      End
      Begin VB.TextBox txtDisplayNameSI 
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
         Left            =   1500
         TabIndex        =   8
         Top             =   3990
         Width           =   1845
      End
      Begin VB.CommandButton cmdAddSave 
         Caption         =   "Add"
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
         Height          =   435
         Left            =   2100
         TabIndex        =   9
         Top             =   5460
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
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
         Height          =   435
         Left            =   810
         TabIndex        =   10
         Top             =   5460
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtUnit 
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
         Left            =   1230
         TabIndex        =   4
         Top             =   1380
         Width           =   2145
      End
      Begin VB.TextBox txtLowRange 
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
         Left            =   1260
         TabIndex        =   2
         Top             =   600
         Width           =   2145
      End
      Begin VB.TextBox txtExamID 
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
         Left            =   1230
         MaxLength       =   8
         TabIndex        =   1
         Top             =   210
         Width           =   555
      End
      Begin VB.TextBox txtHighRange 
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
         Left            =   1230
         TabIndex        =   3
         Top             =   990
         Width           =   2145
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Range SI (Upper):"
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
         Index           =   8
         Left            =   90
         TabIndex        =   34
         Top             =   3240
         Width           =   1290
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Range SI (Lower):"
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
         Index           =   7
         Left            =   90
         TabIndex        =   33
         Top             =   2850
         Width           =   1335
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit SI:"
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
         Index           =   9
         Left            =   870
         TabIndex        =   32
         Top             =   3630
         Width           =   495
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Display SI:"
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
         Index           =   10
         Left            =   630
         TabIndex        =   31
         Top             =   4020
         Width           =   750
      End
      Begin VB.Label txtExamName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   1800
         TabIndex        =   26
         Top             =   210
         Width           =   1575
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Display:"
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
         Index           =   6
         Left            =   630
         TabIndex        =   24
         Top             =   1800
         Width           =   570
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit:"
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
         Index           =   5
         Left            =   870
         TabIndex        =   23
         Top             =   1410
         Width           =   315
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Range (Lower):"
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
         Index           =   3
         Left            =   90
         TabIndex        =   22
         Top             =   630
         Width           =   1155
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Range (Upper):"
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
         Index           =   4
         Left            =   90
         TabIndex        =   21
         Top             =   1020
         Width           =   1110
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Name:"
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
         Left            =   300
         TabIndex        =   20
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   6675
      Index           =   2
      Left            =   3570
      TabIndex        =   15
      Top             =   60
      Width           =   8115
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   6405
         Left            =   60
         TabIndex        =   16
         Top             =   180
         Visible         =   0   'False
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   11298
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
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   3
      Left            =   3570
      TabIndex        =   27
      Top             =   6750
      Width           =   8115
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5460
         TabIndex        =   30
         Top             =   180
         Width           =   1275
      End
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
         Height          =   345
         Left            =   6780
         TabIndex        =   28
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Double Click an Item to Edit."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   120
         TabIndex        =   29
         Top             =   210
         Visible         =   0   'False
         Width           =   2280
      End
   End
End
Attribute VB_Name = "frmNormalValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim yClassForms As New ClassForms
Dim rstData As ADODB.Recordset

Private Sub cmdAddSave_Click()
If Len(Trim(txtLowRange.Text)) = 0 Then txtLowRange.Text = "0"
If Len(Trim(txtHighRange.Text)) = 0 Then txtHighRange.Text = "0"
If Len(Trim(txtLowRangeSI.Text)) = 0 Then txtLowRangeSI.Text = "0"
If Len(Trim(txtHighRangeSI.Text)) = 0 Then txtHighRangeSI.Text = "0"
If cmdAddSave.Caption = "Add" Then
    If Trim(txtHighRange.Text) = "" Or Trim(txtLowRange) = "" Or Not IsNumeric(txtLowRange.Text) Or Not IsNumeric(txtHighRange.Text) Then
        MsgBox "Please check the values", vbOKOnly + vbInformation, "Normal Values"
        txtHighRange.SetFocus
        highLight
    Else
        myConnection.Execute ("Insert into tblabclassexams (InstrumentID, ClassID, ExamID, RangeUp, RangeDown, RangeUnit, Display, Active, RangeUpSI, RangeDownSI, RangeUnitSI, DisplaySI )" & _
            "Select '" & InstrumentID & "','" & Replace(txtClassID.Text, "'", "''") & "','" & txtExamID.Text & "'," & txtHighRange.Text & "," & txtLowRange.Text & ",'" & Replace(txtUnit.Text, "'", "''") & "','" & Replace(txtDisplayName.Text, "'", "''") & "',TRUE," & txtHighRangeSI.Text & "," & txtLowRangeSI.Text & ",'" & Replace(txtUnitSI.Text, "'", "''") & "','" & Replace(txtDisplayNameSI.Text, "'", "''") & "'")
        displayExams
        clearItems
        txtClassID.SetFocus
    End If
Else
    myConnection.Execute ("Update tbLabClassExams set RangeUP =" & txtHighRange.Text & ",RangeDown = " & txtLowRange.Text & "," & _
        "RangeUnit = '" & txtUnit.Text & "',Display = '" & txtDisplayName.Text & "', " & "RangeUPSI =" & txtHighRangeSI.Text & ",RangeDownSI = " & txtLowRangeSI.Text & ", RangeUnitSI = '" & txtUnitSI.Text & "',DisplaySI = '" & txtDisplayNameSI.Text & "' " & "where InstrumentID = '" & InstrumentID & "' and ClassID = '" & txtClassID.Text & "' and ExamID = '" & msfgResult.TextMatrix(msfgResult.Row, 11) & "'")
    fraBody(0).Enabled = True
    fraBody(2).Enabled = True
    fraBody(1).Enabled = True
    cmdAddSave.Caption = "Add"
    cmdCancel.Visible = False
    txtExamID.Locked = False
    displayExams
    clearItems
End If
End Sub

Private Sub cmdCancel_Click()
    fraBody(0).Enabled = True
    fraBody(2).Enabled = True
    fraBody(1).Enabled = True
    cmdAddSave.Caption = "Add"
    cmdCancel.Visible = False
    txtExamID.SetFocus
    clearItems
    txtExamID.Locked = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopyFrom_Click()
Dim ctr As Integer
Dim totalcount As Integer
Dim rstCopy As ADODB.Recordset
Dim rstTrans As ADODB.Recordset
yClassForms.showSearchCode 2, "", , "Select Classification to copy from..."
Set rstCopy = myConnection.Execute("Select * from tbLabClassExams where classid = '" & yClassForms.LastsearchID & "' and InstrumentID = '" & InstrumentID & "'")
If rstCopy.RecordCount > 0 Then
    rstCopy.MoveFirst
    For ctr = 1 To rstCopy.RecordCount
        Set rstTrans = myConnection.Execute("Select * from tbLabClassExams where ClassID = '" & txtClassID.Text & "' and InstrumentID = '" & InstrumentID & "' and ExamID = '" & rstCopy.Fields("ExamID").Value & "'")
        If rstTrans.RecordCount = 0 Then
            myConnection.Execute "Insert into tblabclassexams (InstrumentID, ClassID, ExamID, RangeUp, RangeDown, RangeUnit, Display, Active, RangeUpSI, RangeDownSI, RangeUnitSI, DisplaySI )" & _
                        "Select '" & InstrumentID & "','" & txtClassID.Text & "','" & rstCopy.Fields("ExamID").Value & "'," & rstCopy.Fields("RangeUP").Value & "," & rstCopy.Fields("RangeDown").Value & ",'" & rstCopy.Fields("RangeUnit").Value & "','" & rstCopy.Fields("Display").Value & "'," & rstCopy.Fields("Active").Value & "," & rstCopy.Fields("RangeUPSI").Value & "," & rstCopy.Fields("RangeDownSI").Value & ",'" & rstCopy.Fields("RangeUnitSI").Value & "','" & rstCopy.Fields("DisplaySI").Value & "'"
                        '"Select '" & InstrumentID & "','" & txtClassID.Text & "','" & txtExamID.Text & "'," & txtHighRange.Text & "," & txtLowRange.Text & ",'" & txtUnit.Text & "','" & txtDisplayName.Text & "',TRUE," & txtHighRangeSI.Text & "," & txtLowRangeSI.Text & ",'" & txtUnitSI.Text & "','" & txtDisplayNameSI.Text & "'")
        End If
        rstCopy.MoveNext
    Next ctr
    txtClassName_Change
Else
    MsgBox "No record to copy...", , App.ProductName
End If
End Sub

Private Sub cmdDeleteItems_Click()
If MsgBox("Are you sure you want to delete all items in this classification?", vbYesNo + vbQuestion + vbDefaultButton2, App.ProductName) = vbYes Then
    myConnection.Execute "Delete from tblabclassexams where classid = '" & txtClassID.Text & "'"
    txtClassName_Change
End If
End Sub

Private Sub cmdPrint_Click()
    If Len(Trim(txtClassName.Caption)) > 0 Then
        If UCase(Dir(App.Path & "\Reports", vbDirectory)) = "REPORTS" Then
            OpenMainReport App.Path & "\reports\normalvalues_01.rpt", InstrumentID, txtClassID.Text, txtClassName.Caption
        ElseIf UCase(Dir(Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports", vbDirectory)) = "REPORTS" Then
            OpenMainReport Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports\normalvalues_01.rpt", InstrumentID, txtClassID.Text, txtClassName.Caption
        End If
        ShowReportViewer False, "Result Print Preview", False
    End If
End Sub

Private Sub Form_Load()
setFormColors Me
    txtInstrumentID.Text = InstrumentID
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rstData = rstData
End Sub

Private Sub msfgResult_DblClick()
    fraBody(0).Enabled = False
    fraBody(2).Enabled = False
    fraBody(1).Enabled = True
    txtExamID.SetFocus
    cmdAddSave.Caption = "Update"
    cmdCancel.Visible = True
    SetDetails
    txtExamID.Locked = True
    cmdAddSave.Enabled = True
End Sub

Private Sub msfgResult_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete And Shift = 1 Then
    myConnection.Execute ("Update tbLabClassExams set Active = FALSE where InstrumentID = '" & InstrumentID & "' and ClassID = '" & txtClassID.Text & "' and ExamID = '" & msfgResult.TextMatrix(msfgResult.Row, 11) & "'")
    displayExams
End If
If KeyCode = vbKeyInsert And Shift = 1 Then
    myConnection.Execute ("Update tbLabClassExams set Active = TRUE where InstrumentID = '" & InstrumentID & "' and ClassID = '" & txtClassID.Text & "' and ExamID = '" & msfgResult.TextMatrix(msfgResult.Row, 11) & "'")
    displayExams
End If
End Sub

Private Sub txtClassID_Change()
    txtClassName.Caption = ""
    fraBody(1).Enabled = False
    msfgResult.Visible = False
    lblStatus.Visible = False
    clearItems
End Sub

Private Sub txtClassID_GotFocus()
    highLight
End Sub

Private Sub txtClassID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    yClassForms.showSearchCode 2, txtClassID.Text
    txtClassID.Text = yClassForms.LastsearchID
    txtClassName.Caption = yClassForms.LastSearchName
    If Len(Trim(txtClassName.Caption)) > 0 Then
        displayExams
'        cmdAddSave.Enabled = True
    Else
'        cmdAddSave.Enabled = False
    End If
End If
End Sub

Private Sub txtClassID_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtClassName_Change()
    If Len(Trim(txtClassName.Caption)) > 0 Then
        displayExams
'        cmdAddSave.Enabled = True
        fraBody(1).Enabled = True
    Else
'        cmdAddSave.Enabled = False
        fraBody(1).Enabled = False
    End If
End Sub

Private Sub txtDisplayName_GotFocus()
    highLight
End Sub

Private Sub txtDisplayName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtDisplayNameSI_GotFocus()
    highLight
End Sub

Private Sub txtDisplayNameSI_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtExamID_Change()
    txtExamName.Caption = ""
    cmdAddSave.Enabled = False
    txtHighRange.Text = ""
    txtLowRange.Text = ""
    txtUnit.Text = ""
    txtDisplayName.Text = ""
End Sub

Private Sub txtExamID_GotFocus()
highLight
End Sub

Private Sub txtExamID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    yClassForms.showSearchCode 3, txtExamID.Text
    txtExamID.Text = yClassForms.LastsearchID
    txtExamName.Caption = yClassForms.LastSearchName
End If
End Sub

Private Sub txtExamID_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtExamName_Change()
If cmdAddSave.Caption = "Add" Then
    If rstData.RecordCount > 0 Then
        rstData.MoveFirst
        rstData.Find "ExamID = '" & txtExamID.Text & "'"
        If rstData.EOF Then
            cmdAddSave.Enabled = True
        Else
            MsgBox "Exam Already exists!", vbOKOnly + vbInformation, "Exam Check"
            txtExamID.SetFocus
            highLight
            cmdAddSave.Enabled = False
        End If
    Else
        cmdAddSave.Enabled = True
    End If
'    If Len(Trim(cmdAddSave.Caption)) = 0 Then cmdAddSave.Enabled = False
End If
End Sub

Private Sub txtHighRange_Change()
    txtDisplayName.Text = (txtLowRange.Text & " - " & txtHighRange.Text)
End Sub

Private Sub txtHighRange_GotFocus()
    highLight
End Sub

Private Sub txtHighRange_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtHighRange_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(1, "1234567890-.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtHighRangeSI_Change()
    txtDisplayNameSI.Text = txtLowRangeSI.Text & " - " & txtHighRangeSI.Text
End Sub

Private Sub txtHighRangeSI_GotFocus()
    highLight
End Sub

Private Sub txtHighRangeSI_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtHighRangeSI_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(1, "1234567890-.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtLowRange_Change()
     txtDisplayName.Text = txtLowRange.Text & " - " & txtHighRange.Text
End Sub

Private Sub txtLowRange_GotFocus()
    highLight
End Sub

Private Sub txtLowRange_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtLowRange_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(1, "1234567890-.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtLowRangeSI_Change()
    txtDisplayNameSI.Text = txtLowRangeSI.Text & " - " & txtHighRangeSI.Text
End Sub

Private Sub txtLowRangeSI_GotFocus()
    highLight
End Sub

Private Sub txtLowRangeSI_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtLowRangeSI_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(1, "1234567890-.", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtUnit_Change()
    'txtDisplayName.Text = txtLowRange.Text & " - " & txtHighRange.Text & " " & txtUnit.Text
End Sub

Private Sub txtUnit_GotFocus()
    highLight
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub clearItems()
txtExamID.Text = ""
txtHighRange.Text = ""
txtLowRange.Text = ""
txtUnit.Text = ""
txtDisplayName.Text = ""
txtHighRangeSI.Text = ""
txtLowRangeSI.Text = ""
txtUnitSI.Text = ""
txtDisplayNameSI.Text = ""
End Sub

Private Sub displayExams()
setGrid
With msfgResult
    Set rstData = myConnection.Execute("sp_DisplayClassItems '" & InstrumentID & "','" & txtClassID.Text & "'")
    If rstData.RecordCount > 0 Then
        Set .DataSource = rstData
        lblStatus.Visible = True
        .Visible = True
    Else
        lblStatus.Visible = False
        .Visible = False
    End If
End With
End Sub

Private Sub setGrid()
With msfgResult
    .ColWidth(0) = .Width * 0.05 - (5 * 3)
    .ColWidth(1) = .Width * 0.25 - (25 * 3)
    .ColWidth(2) = .Width * 0.15 - (20 * 3)
    .ColWidth(3) = .Width * 0.15 - (20 * 3)
    .ColWidth(4) = .Width * 0.1 - (15 * 3)
    .ColWidth(5) = .Width * 0.1 - (15 * 3)
    .ColWidth(6) = .Width * 0.2 - (20 * 3)
    .ColWidth(7) = .Width * 0.2 - (20 * 3)
    .ColWidth(8) = .Width * 0.15 - (15 * 3)
    .ColWidth(9) = .Width * 0.15 - (15 * 3)
    .ColWidth(10) = .Width * 0.1 - (15 * 3)

'    .ColWidth(7) = 0
'    .ColWidth(8) = 0
'    .ColWidth(9) = 0
'    .ColWidth(10) = 0
End With
End Sub

Private Sub SetDetails()
With msfgResult
    txtExamID.Text = .TextMatrix(.Row, 11)
    txtHighRange.Text = Trim(.TextMatrix(.Row, 3))
    txtLowRange.Text = Trim(.TextMatrix(.Row, 2))
    txtUnit.Text = Trim(.TextMatrix(.Row, 4))
    txtDisplayName.Text = Trim(.TextMatrix(.Row, 5))
    
    txtHighRangeSI.Text = Trim(.TextMatrix(.Row, 7))
    txtLowRangeSI.Text = Trim(.TextMatrix(.Row, 6))
    txtUnitSI.Text = Trim(.TextMatrix(.Row, 8))
    txtDisplayNameSI.Text = Trim(.TextMatrix(.Row, 9))

    txtExamName.Caption = yClassForms.showSearchCode(3, txtExamID.Text)
    
End With
End Sub

Private Sub txtUnitSI_GotFocus()
    highLight
End Sub

Private Sub txtUnitSI_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub
