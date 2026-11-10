VERSION 5.00
Begin VB.Form frmTemplate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Templates"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3915
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   3915
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5100
      TabIndex        =   0
      Top             =   2100
      Width           =   3825
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Templates"
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
         Width           =   1725
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   3795
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   3825
      Begin VB.ListBox lstReportName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   60
         TabIndex        =   6
         Top             =   210
         Width           =   3705
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   705
      Index           =   1
      Left            =   30
      TabIndex        =   3
      Top             =   3870
      Width           =   3825
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
         Left            =   1320
         TabIndex        =   5
         Top             =   210
         Width           =   1185
      End
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
         Left            =   2550
         TabIndex        =   4
         Top             =   210
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    TemplateValue = ""
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    lstReportName_DblClick
End Sub

Private Sub Form_Load()
Dim rstTemp As ADODB.Recordset
setFormColors Me
lstReportName.Clear
Set rstTemp = myConnection.Execute("Select rptName from tbLabReportNames where rptSectionCode = '" & SectionCode & "'")
If rstTemp.RecordCount > 0 Then
    rstTemp.MoveFirst
    Do While Not rstTemp.EOF
        With lstReportName
            .AddItem rstTemp.Fields("rptName").Value & ""
        End With
        rstTemp.MoveNext
    Loop
Else
End If
'With fileTemplate
'    .Path = App.Path & "\Reports"
'    .Pattern = SectionCode & "_*.rpt"
'End With
End Sub

Private Sub lstReportName_DblClick()
    TemplateValue = lstReportName.Text
    Unload Me
End Sub
