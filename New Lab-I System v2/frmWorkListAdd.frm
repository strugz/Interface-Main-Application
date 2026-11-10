VERSION 5.00
Begin VB.Form frmWorkListAdd 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9150
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add to Worklist..."
      Height          =   465
      Left            =   5520
      TabIndex        =   4
      Top             =   5400
      Width           =   1635
   End
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   -90
      Width           =   9135
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Worklist - [Details]"
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
         TabIndex        =   3
         Top             =   150
         Width           =   1845
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   465
      Left            =   7290
      TabIndex        =   1
      Top             =   5400
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   5715
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   9135
   End
End
Attribute VB_Name = "frmWorkListAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub
