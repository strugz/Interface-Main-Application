VERSION 5.00
Begin VB.Form frmBackupList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup List"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
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
      Left            =   2280
      TabIndex        =   5
      Top             =   3330
      Width           =   1305
   End
   Begin VB.ListBox List1 
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
      Height          =   3180
      ItemData        =   "frmBackupList.frx":0000
      Left            =   90
      List            =   "frmBackupList.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   60
      Width           =   3495
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
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
      Left            =   900
      TabIndex        =   3
      Top             =   3330
      Width           =   1305
   End
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4380
      TabIndex        =   0
      Top             =   1980
      Width           =   3735
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Backup List"
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
         Width           =   1095
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   4035
      Index           =   0
      Left            =   4740
      TabIndex        =   2
      Top             =   3480
      Width           =   3735
   End
End
Attribute VB_Name = "frmBackupList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ListSubDirs(Path)
    On Error Resume Next
    Dim Count, D(), i, DirName      ' Declare variables.
    DirName = Dir(Path, 16)         ' Get first directory name.
    Do While DirName <> ""
        ' A file or directory name was returned
        If DirName <> "." And DirName <> ".." Then
            ' Not a parent or current directory entry so process it
            If GetAttr(Path + DirName) = 16 Then
                ' This is a directory
                If (Count Mod 10) = 0 Then
                    ' Resize the array
                    ReDim Preserve D(Count + 10)
                End If
                Count = Count + 1   ' Increment counter.
                D(Count) = DirName  ' Add directory name to array
            End If
        End If
        DirName = Dir   ' Get another directory name.
    Loop
    For i = 1 To Count
        ' Put each directory's name in the list box.
        List1.AddItem D(i)
        ' The following is the recursive call
        ' Find this directory's sub-directories
        ListSubDirs Path & D(i) & "\"
    Next i
    DoEvents
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    If Len(Trim(List1.Text)) > 0 Then strBackupPath = List1.Text
    Unload Me
End Sub

Private Sub Form_Load()
    setFormColors Me
    ListSubDirs Replace(Format(DBDatabase, ">"), "LABORATORY.MDB", "") & "Backup\"
End Sub

