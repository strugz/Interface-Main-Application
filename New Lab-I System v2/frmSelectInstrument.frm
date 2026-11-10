VERSION 5.00
Begin VB.Form frmSelectInstrument 
   BorderStyle     =   0  'None
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3270
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   4635
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Plugins"
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
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   2955
      Left            =   0
      TabIndex        =   2
      Top             =   300
      Width           =   4635
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         Default         =   -1  'True
         Height          =   465
         Left            =   1770
         TabIndex        =   5
         Top             =   2340
         Width           =   1305
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         Height          =   1920
         Left            =   120
         Pattern         =   "Host_*.dll"
         TabIndex        =   4
         Top             =   240
         Width           =   4395
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   465
         Left            =   3180
         TabIndex        =   3
         Top             =   2340
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmSelectInstrument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

