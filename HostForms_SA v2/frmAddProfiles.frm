VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddProfiles 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   30
      TabIndex        =   2
      Top             =   -60
      Width           =   6135
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add Profile"
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
         Top             =   180
         Width           =   1065
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   360
      Width           =   6135
      Begin MSComctlLib.ListView lvwProfiles 
         Height          =   3615
         Left            =   60
         TabIndex        =   1
         Top             =   150
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   6376
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Profile ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Profile Name"
            Object.Width           =   7408
         EndProperty
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   705
      Index           =   1
      Left            =   30
      TabIndex        =   4
      Top             =   4140
      Width           =   6135
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         Default         =   -1  'True
         Height          =   405
         Left            =   3000
         TabIndex        =   6
         Top             =   210
         Width           =   1455
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   405
         Left            =   4470
         TabIndex        =   5
         Top             =   210
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmAddProfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstProfiles As ADODB.Recordset
Private ySelectedItems As String

Private Sub cmdClose_Click()
    ReturnValueStr = ySelectedItems
    Unload Me
End Sub

Private Sub cmdSelect_Click()
Dim tempStr As String
Dim ctr As Integer
Dim hasChecked As Boolean
hasChecked = False
With lvwProfiles
    For ctr = 1 To .ListItems.Count
        If .ListItems(ctr).Checked Then
            tempStr = tempStr & .ListItems(ctr).Text & ","
            hasChecked = True
        End If
    Next ctr
    If hasChecked Then tempStr = Mid(tempStr, 1, Len(tempStr) - 1)
End With
ReturnValueStr = tempStr
Unload Me
End Sub

Private Sub Form_Load()
    setFormColors Me
    Set rstProfiles = myConnection.Execute("Select * from tblabprofiles where instrumentid = '" & InstrumentID & "' and status = '1'")
    setListView
End Sub

Private Sub setListView()
Dim ctr As Integer
Dim ctr2 As Integer
Dim arrStr As Variant
lvwProfiles.ListItems.Clear
arrStr = Split(ySelectedItems, ",")
rstProfiles.MoveFirst
For ctr = 1 To rstProfiles.RecordCount
    With lvwProfiles
        .ListItems.Add 1, , rstProfiles.Fields("ProfileID").Value
        .ListItems.Item(1).ListSubItems.Add , , rstProfiles.Fields("ProfileName").Value
        For ctr2 = 0 To UBound(arrStr)
            If arrStr(ctr2) = rstProfiles.Fields("ProfileID").Value Then .ListItems.Item(ctr).Checked = True
        Next ctr2
    End With
    rstProfiles.MoveNext
Next ctr
End Sub

Public Property Let wSelectedItems(ByVal vNewValue As String)
        ySelectedItems = vNewValue
End Property
