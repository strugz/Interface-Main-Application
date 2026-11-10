VERSION 5.00
Begin VB.Form frmSelectPlugins 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Plug-Ins"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Default         =   -1  'True
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
      Height          =   495
      Left            =   2010
      TabIndex        =   2
      Top             =   4380
      Width           =   1215
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
      Height          =   495
      Left            =   3300
      TabIndex        =   1
      Top             =   4380
      Width           =   1215
   End
   Begin VB.FileListBox File1 
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
      Height          =   4230
      Left            =   150
      Pattern         =   "HOST_*.dll"
      TabIndex        =   0
      Top             =   60
      Width           =   4395
   End
End
Attribute VB_Name = "frmSelectPlugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mDef As Object
Private mDefOrg As Object

Private Sub cmdClose_Click()
    returnValue = ""
    Unload Me
End Sub

Private Sub cmdSelect_Click()
returnValue = File1.FileName
Unload Me
End Sub

Private Sub File1_Click()
Dim tempStr As String
If Len(Trim(File1.FileName)) > 0 Then
    tempStr = Format(File1.FileName, ">")
    tempStr = Replace(tempStr, ".DLL", "")
    Set mDef = CreateObject(tempStr & ".MessageDef")
    If Val(GetSetting("Lab-I", "Settings", "EnablePlugin", "0")) = 1 Then
        tempStr = Format(GetSetting("Lab-I", "Settings", "PlugIN", ""), ">")
        Set mDefOrg = CreateObject(tempStr & ".MessageDef")
        If mDef.instrumentid = mDefOrg.instrumentid Then cmdSelect.Enabled = True Else cmdSelect.Enabled = False
    Else
        If mDef.instrumentid = "HMX" Then cmdSelect.Enabled = True Else cmdSelect.Enabled = False
    End If
End If
End Sub

Private Sub Form_Load()
File1.Path = App.Path & "\Plugins"
returnValue = ""
End Sub
