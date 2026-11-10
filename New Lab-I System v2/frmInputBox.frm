VERSION 5.00
Begin VB.Form frmInputBox 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Title"
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5040
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   1590
      Width           =   4935
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4650
         MouseIcon       =   "frmInputBox.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   180
         Width           =   135
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   180
         Width           =   360
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   4560
         Top             =   150
         Width           =   285
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   795
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4935
      Begin VB.TextBox txtValue 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   90
         TabIndex        =   3
         Top             =   240
         Width           =   4755
      End
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private yValue As String
Private yTitle As String
Private yTop As Integer
Private yLeft As Integer
Private yMaxLength As Integer
Private yUpcase As Boolean
Private yMask As Boolean

Public Property Let wMask(ByVal vNewValue As Integer)
    yMask = vNewValue
End Property

Public Property Let wTop(ByVal vNewValue As Integer)
    yTop = vNewValue
End Property

Public Property Let wLeft(ByVal vNewValue As Integer)
    yLeft = vNewValue
End Property

Public Property Let wTitle(ByVal vNewValue As String)
    yTitle = vNewValue
End Property

Private Sub Form_Load()
    setFormColors Me
    txtValue.Text = yValue
    If yMask Then txtValue.PasswordChar = "•"
    Me.Top = yTop
    Me.Left = yLeft
    Me.lblHeader(0).Caption = yTitle
    txtValue.MaxLength = yMaxLength
End Sub

Private Sub Label2_Click()
    frmSearchVal = ""
    Unload Me
End Sub

Private Sub lblHeader_Change(Index As Integer)
    Me.Caption = lblHeader(0).Caption
End Sub

Private Sub txtValue_GotFocus()
    Highlight
End Sub

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    frmSearchVal = txtValue.Text
    Unload Me
ElseIf KeyCode = vbKeyEscape Then
    frmSearchVal = ""
    Unload Me
End If
End Sub

Public Property Let wValue(ByVal vNewValue As String)
    yValue = vNewValue
End Property

Public Property Let wMaxLength(ByVal vNewValue As Integer)
    yMaxLength = vNewValue
End Property

Public Property Let wUpcase(ByVal vNewValue As Boolean)
    yUpcase = vNewValue
End Property

Private Sub txtValue_KeyPress(KeyAscii As Integer)
If yUpcase = True Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
