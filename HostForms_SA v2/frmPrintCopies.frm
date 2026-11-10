VERSION 5.00
Begin VB.Form frmPrintCopies 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Result Printout"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5670
      TabIndex        =   6
      Top             =   2040
      Width           =   3345
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result Printout"
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
         TabIndex        =   7
         Top             =   150
         Width           =   1470
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   2625
      Index           =   0
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   4245
      Begin VB.CheckBox chkPrinted 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2340
         TabIndex        =   12
         Top             =   600
         Width           =   195
      End
      Begin VB.TextBox txtCopy2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1590
         TabIndex        =   3
         Top             =   2040
         Width           =   2595
      End
      Begin VB.TextBox txtCopy2Tag 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   1710
         Width           =   4005
      End
      Begin VB.TextBox txtCopy1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1590
         TabIndex        =   0
         Top             =   600
         Width           =   675
      End
      Begin VB.CheckBox chkCopy2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   180
         TabIndex        =   1
         Top             =   1410
         Width           =   195
      End
      Begin VB.Label lbltext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   " [F3] - With Tag"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   5
         Left            =   390
         TabIndex        =   16
         Top             =   1410
         Width           =   3675
      End
      Begin VB.Label lbltext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Mark as printed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   4
         Left            =   2580
         TabIndex        =   15
         Top             =   630
         Width           =   1350
      End
      Begin VB.Label lbltext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Secondary Copy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   180
         TabIndex        =   14
         Top             =   1080
         Width           =   1920
      End
      Begin VB.Label lbltext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Main Copy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   13
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label lbltext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Copies"
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
         Left            =   210
         TabIndex        =   11
         Top             =   2070
         Width           =   1290
      End
      Begin VB.Label lbltext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Copies"
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
         Left            =   150
         TabIndex        =   9
         Top             =   630
         Width           =   1290
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   675
      Index           =   1
      Left            =   60
      TabIndex        =   10
      Top             =   2550
      Width           =   4245
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
         Height          =   375
         Left            =   2940
         TabIndex        =   5
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print..."
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
         Height          =   375
         Left            =   1650
         TabIndex        =   4
         Top             =   180
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPrintCopies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkCopy2_Click()
    If chkCopy2.Value = 1 Then
        txtCopy2.Enabled = True
        txtCopy2Tag.Enabled = True
    Else
        txtCopy2.Enabled = False
        txtCopy2Tag.Enabled = False
    End If
End Sub

Private Sub cmdClose_Click()
    ReturnValueStr = "0|0|0|0"
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    SaveSetting "HostForms_SA", "Print", "Copy1", Trim(txtCopy1.Text)
    SaveSetting "HostForms_SA", "Print", "Copy2", Trim(txtCopy2.Text)
    SaveSetting "HostForms_SA", "Print", "Copy2Tag", Trim(txtCopy2Tag.Text)
    SaveSetting "HostForms_SA", "Print", "MarkPrinted", Val(chkPrinted.Value)
    ReturnValueStr = Trim(txtCopy1.Text) & "|" & Trim(chkCopy2.Value) & "|" & txtCopy2.Text & "|" & txtCopy2Tag.Text
    Unload Me
End Sub

Private Sub Form_Load()
    setFormColors Me
    chkCopy2.BackColor = fraBody(0).BackColor
    chkCopy2.ForeColor = lbltext(0).ForeColor
    chkPrinted.BackColor = fraBody(0).BackColor
    chkPrinted.ForeColor = lbltext(0).ForeColor
    txtCopy1.Text = GetSetting("HostForms_SA", "Print", "Copy1", "1")
    txtCopy2.Text = GetSetting("HostForms_SA", "Print", "Copy2", "1")
    txtCopy2Tag.Text = GetSetting("HostForms_SA", "Print", "Copy2Tag", "Duplicate")
    chkPrinted.Value = Val(GetSetting("HostForms_SA", "Print", "MarkPrinted", "0"))
End Sub

Private Sub lblText_Click(Index As Integer)
    If Index = 4 Then
        If chkPrinted.Value = 1 Then chkPrinted.Value = 0 Else chkPrinted.Value = 1
    End If
    If Index = 5 Then
        If chkCopy2.Value = 1 Then chkCopy2.Value = 0 Else chkCopy2.Value = 1
    End If
End Sub

Private Sub txtCopy1_Change()
If Val(txtCopy1.Text) > 20 Then
    MsgBox "Sorry... You can not print more than 20 copies.", vbInformation + vbOKOnly, "Print Error"
    txtCopy1.Text = "1"
    highLight
End If
If Val(txtCopy1.Text) = 0 Then
    chkPrinted.Enabled = True
Else
    chkPrinted.Enabled = False
End If
End Sub

Private Sub txtCopy1_GotFocus()
    highLight
End Sub

Private Sub txtCopy1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
    If chkCopy2.Value = 1 Then
        chkCopy2.Value = 0
    Else
        chkCopy2.Value = 1
        txtCopy2.SetFocus
    End If
End If
End Sub

Private Sub txtCopy1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(1, "1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCopy2_Change()
If Val(txtCopy2.Text) > 20 Then
    MsgBox "Sorry... You can not print more than 20 copies.", vbInformation + vbOKOnly, "Print Error"
    txtCopy2.Text = "1"
    highLight
End If
End Sub

Private Sub txtCopy2_GotFocus()
    highLight
End Sub

Private Sub txtCopy2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(1, "1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCopy2Tag_GotFocus()
    highLight
End Sub

Private Sub txtCopy2Tag_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "|" Then
        KeyAscii = 0
    End If
End Sub
