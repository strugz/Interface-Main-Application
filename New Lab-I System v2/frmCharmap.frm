VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MsHFlxGd.ocx"
Begin VB.Form frmCharmap 
   BackColor       =   &H006B4B01&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Special Characters"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   13.5
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
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
      Height          =   345
      Left            =   1800
      TabIndex        =   2
      Top             =   3960
      Width           =   1095
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
      Height          =   345
      Left            =   660
      TabIndex        =   1
      Top             =   3960
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
      Height          =   3855
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   6800
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmCharmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const charname As String = "NUL,SOH,STX,ETX,EOT,ENQ,ACK,BEL,BS,TAB,LF,VT,FF,CR,SO,SI,DLE,DC1,DC2,DC3,DC4,NAK,SYN,ETB,CAN,EM,SUB,ESC,FS,GS,RS,US,SPACE"

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdInsert_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim ctr As Integer
Dim arrName As Variant
arrName = Split(charname, ",")
ReturnValue = ""
With msfgResult
    .Rows = 34
    .TextMatrix(0, 0) = "NAME"
    .TextMatrix(0, 1) = "CHAR"
    For ctr = 0 To 32
        .TextMatrix(ctr + 1, 1) = Chr(ctr)
        .TextMatrix(ctr + 1, 0) = arrName(ctr)
    Next ctr
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReturnValue = Chr(msfgResult.Row)
End Sub

Private Sub msfgResult_Click()
    cmdInsert_Click
End Sub
