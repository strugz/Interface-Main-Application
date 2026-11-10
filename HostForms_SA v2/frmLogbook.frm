VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLogbook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logbook Report Printout"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3420
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4380
      TabIndex        =   5
      Top             =   2550
      Width           =   3345
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Logbook Report Printout"
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
         TabIndex        =   6
         Top             =   150
         Width           =   2400
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   1185
      Index           =   0
      Left            =   30
      TabIndex        =   7
      Top             =   30
      Width           =   3345
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
      Begin VB.Label lbltext 
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
         TabIndex        =   9
         Top             =   720
         Width           =   690
      End
      Begin VB.Label lbltext 
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
         TabIndex        =   8
         Top             =   330
         Width           =   765
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   675
      Index           =   1
      Left            =   30
      TabIndex        =   4
      Top             =   1200
      Width           =   3345
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print..."
         Height          =   375
         Left            =   780
         TabIndex        =   2
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   180
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmLogbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
'Dim rstTemp As ADODB.Recordset
'Dim ctr As Long
'Dim ctrStart As Long
'Dim ctrEnd As Long
Me.MousePointer = 11
'    myConnection.Execute "Delete from tbsequence"
'    Set rstTemp = myConnection.Execute("SELECT SpecimenID " & _
'            "From tbPatientMaster " & _
'            "WHERE RefNum <> 'HEMA_TESTPREVIEW' AND ResultDate >= CDate('" & Format(DTPicker1.Value, "mm/dd/yyyy HH:nn:ss") & "') And ResultDate<=CDate('" & Format(DTPicker2.Value, "mm/dd/yyyy HH:nn:ss") & "') " & _
'            "AND Printed = True and InstrumentID = '" & InstrumentID & "' and specimenid <> '' order by val(SpecimenID)")
'    If rstTemp.RecordCount > 0 Then
'        rstTemp.MoveFirst
'        ctrStart = Val(rstTemp.Fields("SpecimenID").Value)
'        rstTemp.MoveLast
'        ctrEnd = Val(rstTemp.Fields("SpecimenID").Value)
'        myConnection.BeginTrans
'        For ctr = ctrStart To ctrEnd
'            myConnection.Execute ("insert into tbsequence (sequencenumber) select '" & ctr & "'")
'        Next ctr
'        myConnection.CommitTrans
'    End If
    If UCase(Dir(App.Path & "\Reports", vbDirectory)) = "REPORTS" Then
        OpenMainReport App.Path & "\reports\" & "Logbook.rpt", Format(DTPicker1.Value, "mm/dd/yyyy HH:mm"), Format(DTPicker2.Value, "mm/dd/yyyy HH:mm"), CStr("1")
    ElseIf UCase(Dir(Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports", vbDirectory)) = "REPORTS" Then
        OpenMainReport Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports\" & "Logbook.rpt", Format(DTPicker1.Value, "mm/dd/yyyy HH:mm"), Format(DTPicker2.Value, "mm/dd/yyyy HH:mm"), CStr("1")
    End If
    ShowReportViewer False, "Logbook Preview", False
    Unload Me
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
End Sub

