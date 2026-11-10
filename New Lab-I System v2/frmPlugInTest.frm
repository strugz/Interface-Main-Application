VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPlugInTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plugin Test Center"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8940
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   6
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8940
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
      Left            =   10140
      TabIndex        =   0
      Top             =   2340
      Width           =   2115
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plugin Test Center"
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
         Width           =   1770
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
      Height          =   4695
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   30
      Width           =   8895
      Begin VB.CommandButton Command1 
         Caption         =   "Save to Text"
         Height          =   405
         Left            =   5910
         TabIndex        =   12
         Top             =   4230
         Width           =   1365
      End
      Begin VB.ComboBox cboFormatting 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   4260
         Width           =   2025
      End
      Begin VB.CommandButton cmdPreviousRecord 
         Caption         =   "Previous Record"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2250
         TabIndex        =   9
         Top             =   4230
         Width           =   1425
      End
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
         Height          =   405
         Left            =   7320
         TabIndex        =   7
         Top             =   4230
         Width           =   1455
      End
      Begin VB.CommandButton cmdNextRecord 
         Caption         =   "Next Record"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3690
         TabIndex        =   6
         Top             =   4230
         Width           =   1425
      End
      Begin VB.FileListBox fileHosts 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   120
         Pattern         =   "Host_*.dll"
         TabIndex        =   3
         Top             =   450
         Width           =   2085
      End
      Begin RichTextLib.RichTextBox rtbMain 
         Height          =   3735
         Left            =   2220
         TabIndex        =   4
         Top             =   450
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   6588
         _Version        =   393217
         BackColor       =   12632256
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"frmPlugInTest.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCount 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100/100"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   8070
         TabIndex        =   11
         Top             =   180
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Viewer"
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
         Left            =   2220
         TabIndex        =   8
         Top             =   210
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plugins"
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
         Left            =   150
         TabIndex        =   5
         Top             =   240
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmPlugInTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mDef As Object
Private myResultValue As String
Private rstTemp As ADODB.Recordset

Private Sub cboFormatting_Click()
ViewItems
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNextRecord_Click()
If rstTemp.AbsolutePosition <> rstTemp.RecordCount Then
    rstTemp.MoveNext
Else
    rstTemp.MoveLast
End If
ViewItems
lblCount = rstTemp.AbsolutePosition & "/" & rstTemp.RecordCount
End Sub

Private Sub ViewItems()
Dim rstExam2 As ADODB.Recordset
On Error GoTo errFound
Set rstExam2 = myConnection.Execute("select * from tbcolabexam where Active = True and InstrumentID = '" & InstrumentID & "' and HisExamCode <> '0' and HisResultCode <> '0'")
    Select Case cboFormatting.Text
        Case "Raw Data"
            rtbMain.Text = rstTemp.Fields("ResultData").Value
        Case "OBX Format"
            If Len(Trim(fileHosts.filename)) > 0 Then
                rtbMain.Text = mDef.segment_OBX(rstTemp.Fields("ResultData").Value)
            Else
                rtbMain.Text = segment_OBX(rstTemp.Fields("ResultData").Value, rstExam2)
            End If
        Case "OBI Format"
            If Len(Trim(fileHosts.filename)) > 0 Then
                rtbMain.Text = mDef.segment_oBI(rstTemp.Fields("ResultData").Value)
            Else
                rtbMain.Text = get_OtherResultInfo(rstTemp.Fields("ResultData").Value, 1)
            End If
        Case "Text Format"
            If Len(Trim(fileHosts.filename)) > 0 Then
                rtbMain.Text = mDef.textDisplay(rstTemp.Fields("ResultData").Value)
            Else
                rtbMain.Text = DisplayResult2Text(segment_OBX(rstTemp.Fields("ResultData").Value, rstExam2), get_OtherResultInfo(rstTemp.Fields("ResultData").Value))
            End If
        Case "HL7 OBX"
            If Len(Trim(fileHosts.filename)) > 0 Then
                rtbMain.Text = mDef.HL7_OBX(rstTemp.Fields("ResultData").Value)
            Else
                rtbMain.Text = HL7_OBX(rstTemp.Fields("ResultData").Value, rstExam2)
            End If
        Case "HL7 MSH"
            If Len(Trim(fileHosts.filename)) > 0 Then
                rtbMain.Text = mDef.HL7_MSH(rstTemp.Fields("ResultData").Value)
            Else
                rtbMain.Text = HL7_MSH(rstTemp.Fields("ResultData").Value)
            End If
        Case "HL7 PID"
            If Len(Trim(fileHosts.filename)) > 0 Then
                rtbMain.Text = mDef.HL7_PID(rstTemp.Fields("ResultData").Value)
            Else
                rtbMain.Text = HL7_PID(rstTemp.Fields("ResultData").Value)
            End If
        Case "HL7 ORC"
            If Len(Trim(fileHosts.filename)) > 0 Then
                rtbMain.Text = mDef.HL7_ORC(rstTemp.Fields("ResultData").Value)
            Else
                rtbMain.Text = HL7_ORC(rstTemp.Fields("ResultData").Value)
            End If
        Case "HL7 OBR"
            If Len(Trim(fileHosts.filename)) > 0 Then
                rtbMain.Text = mDef.HL7_OBR(rstTemp.Fields("ResultData").Value)
            Else
                rtbMain.Text = HL7_OBR(rstTemp.Fields("ResultData").Value)
            End If
        Case "HL7 ORU"
            If Len(Trim(fileHosts.filename)) > 0 Then
                rtbMain.Text = mDef.HL7_ORU(rstTemp.Fields("ResultData").Value)
            Else
                rtbMain.Text = HL7_ORU(rstTemp.Fields("ResultData").Value, rstExam2)
            End If
    End Select
errFound:
    If Err.Number = 91 Then
        rtbMain.Text = "No Plugins."
    ElseIf Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, , App.ProductName
    End If
End Sub

Private Sub cmdPreviousRecord_Click()
If rstTemp.AbsolutePosition > 1 Then
    rstTemp.MovePrevious
Else
    rstTemp.MoveFirst
End If
ViewItems
lblCount = rstTemp.AbsolutePosition & "/" & rstTemp.RecordCount
End Sub

Private Sub Command1_Click()
    rtbMain.SaveFile App.Path & "\SampleText.text", rtfText
End Sub

Private Sub fileHosts_Click()
Dim tempStr As String
On Error GoTo errFound
If Len(Trim(fileHosts.filename)) > 0 Then
    tempStr = Format(fileHosts.filename, ">")
    tempStr = Replace(tempStr, ".DLL", "")
    Set mDef = CreateObject(tempStr & ".MessageDef")
    ViewItems
End If
errFound:
    If Err.Number <> 0 Then
        rtbMain.Text = Err.Description
    Else
    
    End If
End Sub

Private Sub Form_Load()
setFormColors Me
fileHosts.Path = App.Path & "\Plugins"
Set rstTemp = myConnection.Execute("select * from labunprocessed")
If rstTemp.RecordCount > 0 Then
    myResultValue = rstTemp.Fields("ResultData").Value
    fillcboFormatting
    ViewItems
    lblCount = rstTemp.AbsolutePosition & "/" & rstTemp.RecordCount
Else
    Unload Me
End If
End Sub

Private Sub fillcboFormatting()
With cboFormatting
    .AddItem "Raw Data"
    .AddItem "OBX Format"
    .AddItem "OBI Format"
    .AddItem "Text Format"
    .AddItem "HL7 ORU"
    .AddItem "HL7 MSH"
    .AddItem "HL7 PID"
    .AddItem "HL7 ORC"
    .AddItem "HL7 OBR"
    .AddItem "HL7 OBX"
    .ListIndex = 0
End With
End Sub

