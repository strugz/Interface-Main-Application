VERSION 5.00
Begin VB.Form frmDailyBackup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Daily Backup"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8115
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   180
      Top             =   1170
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Please wait..."
      BeginProperty Font 
         Name            =   "Century"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   30
      TabIndex        =   1
      Top             =   840
      Width           =   8055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Daily backup started."
      BeginProperty Font 
         Name            =   "Century"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   30
      TabIndex        =   0
      Top             =   300
      Width           =   8055
   End
End
Attribute VB_Name = "frmDailyBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Activate()
    DailyBackup
    Unload Me
End Sub

Private Sub Timer1_Timer()
'    Me.Hide
End Sub
