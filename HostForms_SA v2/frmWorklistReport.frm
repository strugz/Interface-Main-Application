VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmWorklistReport 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Worklist Report Printout"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3825
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   3825
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4620
      TabIndex        =   10
      Top             =   1170
      Width           =   3525
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Worklist Report Printout"
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
         TabIndex        =   11
         Top             =   150
         Width           =   2385
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   1125
      Index           =   1
      Left            =   30
      TabIndex        =   12
      Top             =   30
      Width           =   3765
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   1020
         TabIndex        =   1
         Top             =   660
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MM/dd/yyyy HH:mm"
         Format          =   137756675
         CurrentDate     =   38453
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1020
         TabIndex        =   0
         Top             =   300
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MM/dd/yyyy HH:mm"
         Format          =   137756675
         CurrentDate     =   38453
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date:"
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
         Left            =   180
         TabIndex        =   14
         Top             =   330
         Width           =   765
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "End Date:"
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
         Left            =   270
         TabIndex        =   13
         Top             =   720
         Width           =   690
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      Caption         =   "Sorting"
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
      Height          =   1635
      Index           =   2
      Left            =   30
      TabIndex        =   16
      Top             =   1200
      Width           =   3765
      Begin VB.OptionButton optSort 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "by Date"
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
         Height          =   375
         Index           =   5
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1110
         Width           =   1725
      End
      Begin VB.OptionButton optSort 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "by Physician"
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
         Height          =   375
         Index           =   4
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1110
         Width           =   1725
      End
      Begin VB.OptionButton optSort 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "by Profile"
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
         Height          =   375
         Index           =   3
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   690
         Width           =   1725
      End
      Begin VB.OptionButton optSort 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "by Name"
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
         Height          =   375
         Index           =   2
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   270
         Width           =   1725
      End
      Begin VB.OptionButton optSort 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "by Ward"
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
         Height          =   375
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   690
         Width           =   1725
      End
      Begin VB.OptionButton optSort 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "by Specimen ID"
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
         Height          =   360
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   270
         Value           =   -1  'True
         Width           =   1725
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   675
      Index           =   0
      Left            =   30
      TabIndex        =   15
      Top             =   2850
      Width           =   3765
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   375
         Left            =   2460
         TabIndex        =   9
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print..."
         Height          =   375
         Left            =   1170
         TabIndex        =   8
         Top             =   180
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmWorklistReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SortOrder As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
Me.MousePointer = 11
    If UCase(Dir(App.Path & "\Reports", vbDirectory)) = "REPORTS" Then
        OpenMainReport App.Path & "\reports\" & "WL_" & SectionCode & ".rpt", InstrumentID, Format(DTPicker1.Value, "mm/dd/yyyy HH:mm"), Format(DTPicker2.Value, "mm/dd/yyyy HH:mm"), CStr(SortOrder), rstHospital.Fields("Hospitalname").Value, rstHospital.Fields("HospitalAddress1").Value
    ElseIf UCase(Dir(Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports", vbDirectory)) = "REPORTS" Then
        OpenMainReport Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports\" & "WL_" & SectionCode & ".rpt", InstrumentID, Format(DTPicker1.Value, "mm/dd/yyyy HH:mm"), Format(DTPicker2.Value, "mm/dd/yyyy HH:mm"), CStr(SortOrder), rstHospital.Fields("Hospitalname").Value, rstHospital.Fields("HospitalAddress1").Value
    End If
    ShowReportViewer False, "Worklist Preview", False
    SaveSetting "HostForms_SA", "WorklistReport", "Sorting", CStr(SortOrder)
Me.MousePointer = 0
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub Form_Load()
setFormColors Me
    DTPicker1.Value = DateAdd("n", -480, Now)
    DTPicker2.Value = Now
    SortOrder = Val(GetSetting("HostForms_SA", "WorklistReport", "Sorting", "1"))
    optSort(SortOrder - 1).Value = True
    optSort(0).Caption = "by " & Replace(GetSetting("HostForms_SA", "LabelSettings", "SpecimenID", "Specimen ID"), ":", "")
    optSort(1).Caption = "by " & Replace(GetSetting("HostForms_SA", "LabelSettings", "Ward", "Ward"), ":", "")
    optSort(5).Caption = "by " & Replace(GetSetting("HostForms_SA", "LabelSettings", "Resultdate", "Result Date"), ":", "")
    optSort(4).Caption = "by " & Replace(GetSetting("HostForms_SA", "LabelSettings", "Physician", "Physician"), ":", "")
    optSort(0).BackColor = fraBody(0).BackColor
    optSort(1).BackColor = fraBody(0).BackColor
    optSort(2).BackColor = fraBody(0).BackColor
    optSort(3).BackColor = fraBody(0).BackColor
    optSort(4).BackColor = fraBody(0).BackColor
    optSort(5).BackColor = fraBody(0).BackColor
    optSort(0).ForeColor = lbltext(0).ForeColor
    optSort(1).ForeColor = lbltext(0).ForeColor
    optSort(2).ForeColor = lbltext(0).ForeColor
    optSort(3).ForeColor = lbltext(0).ForeColor
    optSort(4).ForeColor = lbltext(0).ForeColor
    optSort(5).ForeColor = lbltext(0).ForeColor
    fraBody(2).ForeColor = lbltext(0).ForeColor
End Sub

Private Sub optSort_Click(Index As Integer)
    SortOrder = Index + 1
End Sub

Private Sub optSort_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub
