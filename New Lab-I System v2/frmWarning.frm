VERSION 5.00
Begin VB.Form frmWarning 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   240
      Top             =   210
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WARNING!!!!!!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3465
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   7095
   End
End
Attribute VB_Name = "frmWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblWarning.Caption = GetSetting("Lab-I", "Others", "WARNINGMESSAGE", "WARNING")
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub
