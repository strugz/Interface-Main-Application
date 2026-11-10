VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmNumBC 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Barcode"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "cancel"
      Height          =   300
      Left            =   1710
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   915
      Width           =   1005
   End
   Begin VB.TextBox txtNumBC 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   225
      MaxLength       =   2
      TabIndex        =   0
      Top             =   225
      Width           =   2070
   End
   Begin MSComCtl2.UpDown UpDown 
      Height          =   330
      Left            =   2295
      TabIndex        =   1
      Top             =   225
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   582
      _Version        =   393216
      BuddyControl    =   "txtNumBC"
      BuddyDispid     =   196610
      OrigLeft        =   2565
      OrigTop         =   210
      OrigRight       =   2820
      OrigBottom      =   570
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
End
Attribute VB_Name = "frmNumBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private xNumCopies As Integer
Private blnEnterPress As Boolean

Private Sub cmdCancel_Click()
    xNumCopies = 0
    Unload Me
End Sub

Private Sub Form_Load()
    txtNumBC.Text = GetSetting("LMS", "Barcode", "Copy", 1)
    txtNumBC.SelLength = Len(txtNumBC.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If blnEnterPress = True Then
        If xNumCopies > 10 Then xNumCopies = 10
        SaveSetting "LMS", "Barcode", "Copy", xNumCopies
        Unload Me
    Else
        xNumCopies = 0
    End If
End Sub

Private Sub txtNumBC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then
        UpDown_UpClick
    ElseIf KeyCode = 40 Then
        UpDown_DownClick
    End If
    txtNumBC.SelLength = Len(txtNumBC.Text)
End Sub

Private Sub txtNumBC_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        KeyAscii = KeyAscii
    ElseIf KeyAscii = 13 Then
        xNumCopies = Val(txtNumBC.Text)
        blnEnterPress = True
        Unload Me
    Else
        KeyAscii = 0
    End If

End Sub

Private Sub UpDown_DownClick()
    If Val(txtNumBC.Text) <> 0 Then
        txtNumBC.Text = Val(txtNumBC.Text) - 1
    End If
    txtNumBC.SelLength = Len(txtNumBC.Text)
End Sub

Private Sub UpDown_UpClick()
    If Val(txtNumBC.Text) <> 10 Then
        txtNumBC.Text = Val(txtNumBC.Text) + 1
    End If
    txtNumBC.SelLength = Len(txtNumBC.Text)
End Sub

Public Property Get NumCopies() As Integer
    NumCopies = xNumCopies
End Property

