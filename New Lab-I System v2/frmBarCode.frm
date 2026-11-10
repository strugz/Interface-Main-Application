VERSION 5.00
Begin VB.Form frmBarCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Barcode"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   435
      Left            =   270
      TabIndex        =   2
      Top             =   870
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Default         =   -1  'True
      Height          =   435
      Left            =   2220
      TabIndex        =   1
      Top             =   870
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   3885
   End
End
Attribute VB_Name = "frmBarCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
HostSA.PrintBarcode "*" & Mod43(Text1.Text) & "*"
'HostSA.PrintBarcode "*" & Text1.Text & "*"
End Sub

Private Sub cmdPrint_GotFocus()
cmdPrint_Click
End Sub

