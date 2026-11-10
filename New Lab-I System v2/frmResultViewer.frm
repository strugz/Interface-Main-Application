VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmResultViewer 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   6090
   ClientLeft      =   4455
   ClientTop       =   900
   ClientWidth     =   4860
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   -90
      Width           =   4845
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result Entry"
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
         Left            =   90
         TabIndex        =   8
         Top             =   150
         Width           =   1200
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   5775
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   300
      Width           =   4845
      Begin VB.CommandButton cmdSave 
         Caption         =   "Update"
         Height          =   375
         Left            =   1830
         TabIndex        =   4
         Top             =   5220
         Width           =   1455
      End
      Begin VB.TextBox txtResult 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   90
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1890
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtSampleNum 
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
         Height          =   285
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   6
         Top             =   210
         Width           =   3675
      End
      Begin VB.TextBox txtDate 
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
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   870
         Width           =   3675
      End
      Begin VB.TextBox txtName 
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
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   540
         Width           =   3675
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3390
         TabIndex        =   5
         Top             =   5220
         Width           =   1365
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   3795
         Left            =   60
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1290
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   6694
         _Version        =   393216
         BackColor       =   16777215
         FixedCols       =   0
         BackColorFixed  =   12615680
         ForeColorFixed  =   16777215
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   0
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
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Specimen #:"
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
         Left            =   210
         TabIndex        =   12
         Top             =   240
         Width           =   885
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result Date:"
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
         Left            =   240
         TabIndex        =   11
         Top             =   900
         Width           =   870
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Name:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   570
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmResultViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private yInstrumentID As String
Private yPatientName As String
Private yDateTime As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Property Let wInstrumentID(ByVal vNewValue As String)
    yInstrumentID = vNewValue
End Property

Public Property Let wPatientName(ByVal vNewValue As String)
    yPatientName = vNewValue
End Property

Public Property Let wDateTime(ByVal vNewValue As String)
    yDateTime = vNewValue
End Property

Private Sub cmdSave_Click()
    UpdateData
End Sub

Private Sub Form_Load()
setFormColors Me
If DBServerType = "LOCAL DATABASE" Then
    Set msfgResult.Recordset = myConnection.Execute("select  ExamName as [Exam Name],cstr(Result) as Result,Unit,accessionnum from tbLabInstrumentResults where Instrumentid = '" & yInstrumentID & "' and Patientname = '" & yPatientName & "' and resultdate = #" & Format(yDateTime, "mm/dd/yyyy HH:mm") & "#")
ElseIf DBServerType = "MSSQL SERVER" Then
    Set msfgResult.Recordset = myConnection.Execute("select ExamName as [Exam Name],convert(varchar(10),Result,101) as Result,Unit,accessionnum from tbLabInstrumentResults where Instrumentid = '" & yInstrumentID & "' and Patientname = '" & yPatientName & "' and resultdate = '" & Format(yDateTime, "mm/dd/yyyy HH:mm") & "'")
Else
    Set msfgResult.Recordset = myConnection.Execute("select  ExamName as 'Exam Name',Result,Unit,accessionnum from tbLabInstrumentResults where Instrumentid = '" & yInstrumentID & "' and Patientname = '" & yPatientName & "' and resultdate = '" & Format(yDateTime, "yyyy-mm-dd HH:mm:dd") & "'")
End If
setmsfgResult
txtName.Text = yPatientName
txtDate.Text = Format(yDateTime, "dddd - mmmm dd, yyyy HH:mm ampm")
txtSampleNum.Text = msfgResult.TextMatrix(msfgResult.Row, 3)
If txtSampleNum.Visible Then txtSampleNum.SetFocus
End Sub

Private Sub setmsfgResult()
With msfgResult
    If .Rows >= 15 Then
        .ColWidth(0) = (.Width - 300) * 0.4
        .ColWidth(1) = (.Width - 300) * 0.3
        .ColWidth(2) = (.Width - 300) * 0.3
    Else
        .ColWidth(0) = (.Width - 100) * 0.4
        .ColWidth(1) = (.Width - 100) * 0.3
        .ColWidth(2) = (.Width - 100) * 0.3
    End If
    .ColWidth(3) = 0
End With
End Sub

Private Sub UpdateData()
Dim rstCheck As ADODB.Recordset
    If DBServerType = "MSSQL SERVER" Then
        myConnection.Execute "Update tbLabInstrumentResults set Accessionnum = '" & txtSampleNum & "' where Instrumentid = '" & yInstrumentID & "' and Patientname = '" & yPatientName & "' and resultdate = '" & Format(yDateTime, "mm/dd/yyyy HH:mm") & "'"
    ElseIf DBServerType = "ORACLE" Then
        Set rstCheck = myConnection.Execute(Replace("select Accessionnum from <schema>tbLabInstrumentResults where Instrumentid = '" & yInstrumentID & "' and to_char(ResultDate,'MM/DD/YYYY HH24:MI') = '" & Format(yDateTime, "mm/dd/yyyy HH:mm") & "' and AccessionNum  = '" & txtSampleNum.Text & "'", "<schema>", DBSchema))
        If rstCheck.RecordCount = 0 Then
            myConnection.Execute Replace("Update <schema>tbLabInstrumentResults set Accessionnum = '" & txtSampleNum & "' where Instrumentid = '" & yInstrumentID & "' and Patientname = '" & yPatientName & "' and to_char(ResultDate,'MM/DD/YYYY HH24:MI') = '" & Format(yDateTime, "mm/dd/yyyy HH:mm") & "'", "<schema>", DBSchema)
            MsgBox "Record Updated!", vbOKOnly + vbInformation, App.ProductName
        Else
            MsgBox "Accession Number already exists!", vbOKOnly + vbInformation, App.ProductName
        End If
    Else
        myConnection.Execute "Update tbLabInstrumentResults set Accessionnum = '" & txtSampleNum & "' where Instrumentid = '" & yInstrumentID & "' and Patientname = '" & yPatientName & "' and resultdate = #" & Format(yDateTime, "mm/dd/yyyy HH:mm") & "#"
    End If
End Sub


Private Sub txtSampleNum_GotFocus()
    Highlight
End Sub
