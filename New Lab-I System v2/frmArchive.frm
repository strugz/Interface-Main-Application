VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmArchive 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Archive Samples"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4440
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3330
      TabIndex        =   0
      Top             =   3900
      Width           =   4305
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Archive Samples"
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
         Width           =   1605
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   1755
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   4305
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
         Height          =   465
         Left            =   2790
         TabIndex        =   4
         Top             =   1140
         Width           =   1305
      End
      Begin VB.CommandButton cmdArchive 
         Caption         =   "Archive"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1440
         TabIndex        =   3
         Top             =   1140
         Width           =   1305
      End
      Begin MSComCtl2.DTPicker dtEnd 
         Height          =   345
         Left            =   1020
         TabIndex        =   7
         Top             =   570
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   609
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
         Format          =   68747265
         CurrentDate     =   38040
      End
      Begin MSComCtl2.DTPicker dtstart 
         Height          =   315
         Left            =   1020
         TabIndex        =   8
         Top             =   210
         Width           =   3075
         _ExtentX        =   5424
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
         Format          =   68747265
         CurrentDate     =   38040
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
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
         TabIndex        =   6
         Top             =   690
         Width           =   645
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
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
         TabIndex        =   5
         Top             =   300
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdArchive_Click()
    ReturnValue = dtstart.Value & "|" & dtEnd.Value
    Unload Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim rstTemp As ADODB.Recordset
    setFormColors Me
    Set rstTemp = myConnection.Execute("Select min(resultdate) as oldest from tblabinstrumentresults")
    If IsNull(rstTemp.Fields("oldest").Value) Then
        dtstart = DateAdd("m", -3, Now)
        dtEnd = Now
    Else
        dtstart = rstTemp.Fields("oldest").Value
        dtEnd = DateAdd("m", -3, Now)
    End If
    dtEnd = DateAdd("m", -3, Now)
End Sub
