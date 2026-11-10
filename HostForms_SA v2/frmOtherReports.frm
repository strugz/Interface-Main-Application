VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmOtherReports 
   BorderStyle     =   0  'None
   Caption         =   "Other Reports"
   ClientHeight    =   6570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox fileList 
      Height          =   4965
      Left            =   120
      Pattern         =   "OR_*.RPT"
      TabIndex        =   10
      Top             =   870
      Width           =   4125
   End
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   7485
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Other Reports"
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
         TabIndex        =   1
         Top             =   150
         Width           =   1395
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   5685
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   300
      Width           =   7485
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   5190
         TabIndex        =   3
         Top             =   990
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
         Format          =   127926275
         CurrentDate     =   38453
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   5190
         TabIndex        =   4
         Top             =   630
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
         Format          =   127926275
         CurrentDate     =   38453
      End
      Begin VB.Label lbltext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Report:"
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
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   1020
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
         Left            =   4350
         TabIndex        =   6
         Top             =   660
         Width           =   765
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
         Left            =   4440
         TabIndex        =   5
         Top             =   1050
         Width           =   690
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   675
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   5880
      Width           =   7485
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   375
         Left            =   6210
         TabIndex        =   9
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   4950
         TabIndex        =   8
         Top             =   180
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmOtherReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    setFormColors Me
    With fileList
        .Pattern = "MISC_*.rpt"
        If UCase(Dir(App.Path & "\Reports", vbDirectory)) = "REPORTS" Then
            .Path = App.Path & "\Reports"
        ElseIf UCase(Dir(Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports", vbDirectory)) = "REPORTS" Then
            .Path = Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports"
        End If
    End With
End Sub
