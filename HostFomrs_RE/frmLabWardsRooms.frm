VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmWardsRooms 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Hospital Wards and Rooms"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   13605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   13605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Frame fraBody 
      BackColor       =   &H00A66B2B&
      Height          =   5550
      Index           =   0
      Left            =   0
      TabIndex        =   23
      Top             =   -90
      Width           =   13605
      Begin VB.Frame fraRooms 
         Appearance      =   0  'Flat
         BackColor       =   &H00A66B2B&
         Caption         =   "Room"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3525
         Left            =   9390
         TabIndex        =   36
         Top             =   1830
         Width           =   4140
         Begin VB.TextBox txtRoomAvailableRate 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2250
            TabIndex        =   17
            ToolTipText     =   "Rate ID can be manually inputed. If many ID should be separated by "","""
            Top             =   1680
            Width           =   1560
         End
         Begin VB.TextBox txtRoomDefaultExamRate 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2250
            TabIndex        =   18
            Top             =   2040
            Width           =   1560
         End
         Begin VB.CheckBox chkRoomActive 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2250
            TabIndex        =   19
            Top             =   2400
            Width           =   195
         End
         Begin VB.CommandButton cmdRoomClose 
            Caption         =   "Close"
            Enabled         =   0   'False
            Height          =   495
            Left            =   2850
            TabIndex        =   21
            Top             =   2850
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtRoomLabCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2250
            TabIndex        =   14
            Top             =   600
            Width           =   1560
         End
         Begin VB.CommandButton cmdUpdateRoom 
            Caption         =   "Save"
            Enabled         =   0   'False
            Height          =   495
            Left            =   1590
            TabIndex        =   20
            Top             =   2850
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtRoomID 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   2250
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   255
            Width           =   540
         End
         Begin VB.TextBox txtRoomHISCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2250
            TabIndex        =   15
            Top             =   960
            Width           =   1560
         End
         Begin VB.TextBox txtRoomNo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2250
            MaxLength       =   10
            TabIndex        =   16
            Top             =   1320
            Width           =   1560
         End
         Begin VB.Label lblText 
            BackStyle       =   0  'Transparent
            Caption         =   "Available Exam Rate ID"
            Height          =   405
            Index           =   9
            Left            =   390
            TabIndex        =   48
            Top             =   1695
            Width           =   1800
         End
         Begin VB.Label lblText 
            BackStyle       =   0  'Transparent
            Caption         =   "Default Exam Rate ID"
            Height          =   285
            Index           =   17
            Left            =   390
            TabIndex        =   47
            Top             =   2070
            Width           =   1740
         End
         Begin VB.Label lbltvwRoomIndex 
            Height          =   345
            Left            =   3225
            TabIndex        =   45
            Top             =   180
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label lblText 
            BackStyle       =   0  'Transparent
            Caption         =   "Active"
            Height          =   285
            Index           =   5
            Left            =   405
            TabIndex        =   44
            Top             =   2400
            Width           =   765
         End
         Begin VB.Label lblText 
            BackStyle       =   0  'Transparent
            Caption         =   "Lab Code"
            Height          =   285
            Index           =   0
            Left            =   405
            TabIndex        =   43
            Top             =   615
            Width           =   795
         End
         Begin VB.Label lblText 
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
            Height          =   285
            Index           =   10
            Left            =   405
            TabIndex        =   40
            Top             =   285
            Width           =   510
         End
         Begin VB.Label lblText 
            BackStyle       =   0  'Transparent
            Caption         =   "HIS Code"
            Height          =   285
            Index           =   11
            Left            =   405
            TabIndex        =   39
            Top             =   990
            Width           =   795
         End
         Begin VB.Label lblText 
            BackStyle       =   0  'Transparent
            Caption         =   "Room No/Name."
            Height          =   285
            Index           =   12
            Left            =   405
            TabIndex        =   38
            Top             =   1350
            Width           =   1425
         End
      End
      Begin VB.Frame fraDept 
         Appearance      =   0  'Flat
         BackColor       =   &H00A66B2B&
         Caption         =   "Department"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1605
         Left            =   5235
         TabIndex        =   32
         Top             =   180
         Width           =   8295
         Begin VB.TextBox txtDeptDesc 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1890
            TabIndex        =   2
            Top             =   945
            Width           =   3675
         End
         Begin VB.CheckBox chkDeptActive 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1890
            TabIndex        =   3
            Top             =   1305
            Width           =   195
         End
         Begin VB.TextBox txtDeptName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1890
            TabIndex        =   1
            Top             =   600
            Width           =   3675
         End
         Begin VB.TextBox txtDeptID 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   1890
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   255
            Width           =   540
         End
         Begin VB.CommandButton cmdUpdateDept 
            Caption         =   "Save"
            Enabled         =   0   'False
            Height          =   495
            Left            =   5730
            TabIndex        =   4
            Top             =   960
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmdDeptClose 
            Caption         =   "Close"
            Enabled         =   0   'False
            Height          =   495
            Left            =   7005
            TabIndex        =   5
            Top             =   960
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblText 
            BackStyle       =   0  'Transparent
            Caption         =   "Description:"
            Height          =   240
            Index           =   16
            Left            =   405
            TabIndex        =   46
            Top             =   945
            Width           =   1395
         End
         Begin VB.Label lbltvwDeptIndex 
            Height          =   345
            Left            =   7365
            TabIndex        =   41
            Top             =   165
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label lblText 
            BackStyle       =   0  'Transparent
            Caption         =   "Active:"
            Height          =   240
            Index           =   15
            Left            =   405
            TabIndex        =   35
            Top             =   1275
            Width           =   1395
         End
         Begin VB.Label lblText 
            BackStyle       =   0  'Transparent
            Caption         =   "Department Name:"
            Height          =   240
            Index           =   14
            Left            =   405
            TabIndex        =   34
            Top             =   600
            Width           =   1395
         End
         Begin VB.Label lblText 
            BackStyle       =   0  'Transparent
            Caption         =   "Department ID:"
            Height          =   285
            Index           =   13
            Left            =   405
            TabIndex        =   33
            Top             =   285
            Width           =   1140
         End
      End
      Begin VB.Frame fraWards 
         Appearance      =   0  'Flat
         BackColor       =   &H00A66B2B&
         Caption         =   "Ward"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3525
         Left            =   5235
         TabIndex        =   24
         Top             =   1830
         Width           =   4140
         Begin VB.TextBox txtWardAvailableRate 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2205
            TabIndex        =   9
            Top             =   1680
            Width           =   1560
         End
         Begin VB.CommandButton cmdWardClose 
            Caption         =   "Close"
            Enabled         =   0   'False
            Height          =   495
            Left            =   2850
            TabIndex        =   13
            Top             =   2850
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtWardID 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   2205
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   255
            Width           =   540
         End
         Begin VB.TextBox txtWardLabCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2205
            MaxLength       =   50
            TabIndex        =   6
            Top             =   600
            Width           =   1560
         End
         Begin VB.TextBox txtWardHISCode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2205
            MaxLength       =   50
            TabIndex        =   7
            Top             =   960
            Width           =   1560
         End
         Begin VB.TextBox txtWardName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2205
            MaxLength       =   80
            TabIndex        =   8
            Top             =   1320
            Width           =   1560
         End
         Begin VB.TextBox txtWardDefaultExamRate 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2205
            TabIndex        =   10
            Top             =   2040
            Width           =   1560
         End
         Begin VB.CommandButton cmdUpdateWard 
            Caption         =   "Save"
            Enabled         =   0   'False
            Height          =   495
            Left            =   1590
            TabIndex        =   12
            Top             =   2850
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chkWardActive 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2205
            TabIndex        =   11
            Top             =   2430
            Width           =   195
         End
         Begin VB.Label lblText 
            BackStyle       =   0  'Transparent
            Caption         =   "Available Exam Rate ID"
            Height          =   285
            Index           =   8
            Left            =   255
            TabIndex        =   49
            Top             =   1710
            Width           =   1890
         End
         Begin VB.Label lbltvwWardIndex 
            Height          =   345
            Left            =   3000
            TabIndex        =   42
            Top             =   180
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label lblText 
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
            Height          =   285
            Index           =   1
            Left            =   255
            TabIndex        =   31
            Top             =   285
            Width           =   510
         End
         Begin VB.Label lblText 
            BackStyle       =   0  'Transparent
            Caption         =   "Lab Code"
            Height          =   285
            Index           =   2
            Left            =   255
            TabIndex        =   30
            Top             =   615
            Width           =   795
         End
         Begin VB.Label lblText 
            BackStyle       =   0  'Transparent
            Caption         =   "HIS Code"
            Height          =   285
            Index           =   3
            Left            =   255
            TabIndex        =   29
            Top             =   990
            Width           =   885
         End
         Begin VB.Label lblText 
            BackStyle       =   0  'Transparent
            Caption         =   "Ward Name"
            Height          =   285
            Index           =   4
            Left            =   255
            TabIndex        =   28
            Top             =   1365
            Width           =   1320
         End
         Begin VB.Label lblText 
            BackStyle       =   0  'Transparent
            Caption         =   "Default Exam Rate ID"
            Height          =   285
            Index           =   6
            Left            =   255
            TabIndex        =   27
            Top             =   2070
            Width           =   1575
         End
         Begin VB.Label lblText 
            BackStyle       =   0  'Transparent
            Caption         =   "Active"
            Height          =   285
            Index           =   7
            Left            =   255
            TabIndex        =   26
            Top             =   2430
            Width           =   765
         End
      End
      Begin ComctlLib.TreeView tvwDeptWardsRooms 
         Height          =   5085
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   8969
         _Version        =   327682
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   5
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin ComctlLib.ImageList ImageList 
         Left            =   150
         Top             =   5130
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483624
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   6
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmLabWardsRooms.frx":0000
               Key             =   "Data"
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmLabWardsRooms.frx":0352
               Key             =   "Ward"
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmLabWardsRooms.frx":08A4
               Key             =   "Room"
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmLabWardsRooms.frx":0DF6
               Key             =   "Close"
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmLabWardsRooms.frx":1348
               Key             =   "Disabled"
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmLabWardsRooms.frx":189A
               Key             =   "Open"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuDept 
      Caption         =   "Department"
      Begin VB.Menu mnuDeptAdd 
         Caption         =   "Add Dept."
      End
      Begin VB.Menu mnuDeptEdit 
         Caption         =   "Edit Dept."
      End
      Begin VB.Menu mnuDeptView 
         Caption         =   "View Dept."
      End
      Begin VB.Menu mnuWardAdd 
         Caption         =   "Add Ward"
      End
      Begin VB.Menu mnuDeptAddRoom 
         Caption         =   "Add Room"
      End
   End
   Begin VB.Menu mnuWard 
      Caption         =   "Ward"
      Begin VB.Menu mnuWardEdit 
         Caption         =   "Edit Ward"
      End
      Begin VB.Menu mnuWardView 
         Caption         =   "View Ward"
      End
      Begin VB.Menu mnuRoomAdd 
         Caption         =   "Add Room"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuRoom 
      Caption         =   "Room"
      Begin VB.Menu mnuEditRoom 
         Caption         =   "Edit Room"
      End
      Begin VB.Menu mnuRoomView 
         Caption         =   "View Room"
      End
   End
End
Attribute VB_Name = "frmWardsRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsDeptWardsRooms As ADODB.Recordset

Private Sub cmdDeptClose_Click()
    txtDeptID.Text = ""
    txtDeptName.Text = ""
    txtDeptDesc.Text = ""
    chkDeptActive.Value = 0
    txtDeptName.Enabled = False
    txtDeptDesc.Enabled = False
    chkDeptActive.Enabled = False
    fraDept.Enabled = False
    tvwDeptWardsRooms.Enabled = True
    lbltvwDeptIndex.Caption = ""
    cmdUpdateDept.Enabled = False
    cmdUpdateDept.Visible = False
    cmdDeptClose.Enabled = False
    cmdDeptClose.Visible = False
    tvwDeptWardsRooms.SetFocus
End Sub

Private Sub cmdRoomClose_Click()
    fraRooms.Enabled = False
    txtDeptID.Text = ""
    txtDeptName.Text = ""
    txtDeptDesc.Text = ""
    chkDeptActive.Value = 0
    txtRoomID.Text = ""
    txtRoomLabCode.Text = ""
    txtRoomHISCode.Text = ""
    txtRoomNo.Text = ""
    txtRoomAvailableRate.Text = ""
    txtRoomDefaultExamRate.Text = ""
    chkRoomActive.Value = 0
    lbltvwRoomIndex.Caption = ""
    cmdUpdateRoom.Caption = "Save"
    cmdUpdateRoom.Enabled = False
    cmdUpdateRoom.Visible = False
    cmdRoomClose.Enabled = False
    cmdRoomClose.Visible = False
    tvwDeptWardsRooms.Enabled = True
    tvwDeptWardsRooms.SetFocus
End Sub

Private Sub cmdUpdateRoom_Click()
    Dim arryRoomAvRates() As String
    If Len(Trim(txtRoomDefaultExamRate.Text)) > 0 Then
        If checkDefRate(txtRoomAvailableRate.Text, txtRoomDefaultExamRate.Text) = False Then
            MsgBox "Defualt Rate ID does not match in available rate, cancel save", vbInformation + vbOKOnly, "LMS"
            Exit Sub
        End If
    End If
    
    If Len(Trim(txtRoomLabCode.Text)) > 0 And Len(Trim(txtRoomNo.Text)) > 0 Then
        If cmdUpdateRoom.Caption = "Save" Then
            Set rsDeptWardsRooms = myConnection.Execute("EXEC sp_Room_Insert '" & txtRoomLabCode.Text & "'," & _
                         "'" & txtRoomHISCode.Text & "'," & Val(txtDeptID.Text) & "," & _
                         "'" & txtRoomNo.Text & "'," & IIf(Len(txtRoomDefaultExamRate.Text) > 0, _
                         txtRoomDefaultExamRate.Text, "0") & "," & Val(chkRoomActive.Value))
                         
            If rsDeptWardsRooms.Fields("InsertStat") = "RE" Then
                MsgBox "Room " & txtRoomNo.Text & " in department " & tvwDeptWardsRooms.SelectedItem.Text & " already exist", vbInformation + vbOKOnly, "LMS"
            Else
                With rsDeptWardsRooms
                    tvwDeptWardsRooms.Nodes.Add tvwDeptWardsRooms.SelectedItem.Key, tvwChild, "R" & .Fields("InsertStat"), txtRoomNo.Text, _
                        IIf(chkRoomActive.Value = 1, "Room", "Disabled")
                    SaveAVRate txtRoomAvailableRate.Text, , .Fields("InsertStat")
                End With
                
                txtRoomID.Text = ""
                txtRoomLabCode.Text = ""
                txtRoomHISCode.Text = ""
                txtRoomNo.Text = ""
                txtRoomAvailableRate.Text = ""
                txtRoomDefaultExamRate.Text = ""
            End If
            
            chkRoomActive.Value = 1
            txtRoomLabCode.SetFocus
        Else
            Set rsDeptWardsRooms = myConnection.Execute("UPDATE tbWardRooms SET LISCode='" & Trim(txtRoomLabCode.Text) & "'" & _
                        ", HisCode = '" & Trim(txtRoomHISCode.Text) & "', RoomNo = '" & Trim(txtRoomNo.Text) & "'" & _
                        ", DefaultExamRate = " & IIf(Len(txtRoomDefaultExamRate.Text) > 0, txtRoomDefaultExamRate.Text, 0) & _
                        ", Active = " & chkRoomActive.Value & " WHERE Department =" & _
                        Mid(tvwDeptWardsRooms.SelectedItem.Parent.Key, 2) & " AND RoomID = " & _
                        txtRoomID.Text)
            
            DeleteAVRate txtRoomAvailableRate.Text, , txtRoomID.Text
            SaveAVRate txtRoomAvailableRate.Text, , txtRoomID.Text
            
            tvwDeptWardsRooms.Nodes.Item(Val(lbltvwRoomIndex.Caption)).Text = Trim(txtRoomNo.Text)
            tvwDeptWardsRooms.Nodes.Item(Val(lbltvwRoomIndex.Caption)).Image = IIf(chkRoomActive.Value = 1, "Room", "Disabled")
            
            cmdRoomClose_Click
        End If
    Else
        MsgBox "Pleas Fill LIS Code, Room No.", vbInformation + vbOKOnly, "LMS"
    End If
End Sub

Private Sub cmdUpdateWard_Click()
    Dim arryWardAvRates() As String
    If Len(Trim(txtWardDefaultExamRate.Text)) > 0 Then
        If checkDefRate(txtWardAvailableRate.Text, txtWardDefaultExamRate.Text) = False Then
            MsgBox "Defualt Rate ID does not match in available rate, cancel save", vbInformation + vbOKOnly, "LMS"
            Exit Sub
        End If
    End If
    
    If tvwDeptWardsRooms.SelectedItem Is Nothing Then
        MsgBox "Please Select a Department", vbInformation + vbOKOnly, "LMS"
    Else
        If Len(Trim(txtWardLabCode.Text)) > 0 And Len(Trim(txtWardName.Text)) > 0 Then
            If cmdUpdateWard.Caption = "Save" Then
                Set rsDeptWardsRooms = myConnection.Execute("EXEC sp_Ward_Insert '" & txtWardLabCode.Text & "'," & _
                                    "'" & txtWardHISCode.Text & "','" & txtWardName.Text & "'," & _
                                    "'" & Mid(tvwDeptWardsRooms.SelectedItem.Key, 2) & "','" & txtWardDefaultExamRate.Text & "'," & _
                                    Val(chkWardActive.Value))
                If rsDeptWardsRooms.Fields("InsertStat") = "RE" Then
                    MsgBox "Ward " & txtWardName.Text & " in department of " & tvwDeptWardsRooms.SelectedItem.Text & " already exist", vbInformation + vbOKOnly, "LMS"
                Else
                    With rsDeptWardsRooms
                        tvwDeptWardsRooms.Nodes.Add tvwDeptWardsRooms.SelectedItem.Key, tvwChild, "W" & .Fields("InsertStat"), txtWardName.Text, "Ward"
                        SaveAVRate txtWardAvailableRate.Text, .Fields("InsertStat").Value
                    End With
                    txtWardID.Text = ""
                    txtWardLabCode.Text = ""
                    txtWardHISCode.Text = ""
                    txtWardName.Text = ""
                    txtWardAvailableRate.Text = ""
                    txtWardDefaultExamRate.Text = ""
                End If
                
                chkWardActive.Value = 1
                txtWardLabCode.SetFocus
            Else
                myConnection.Execute ("EXEC sp_Ward_Update " & txtWardID.Text & "," & _
                                    "'" & txtWardLabCode.Text & "'," & _
                                    "'" & txtWardHISCode.Text & "','" & txtWardName.Text & "'," & _
                                    "'" & Mid(tvwDeptWardsRooms.SelectedItem.Parent.Key, 2) & "','" & _
                                    txtWardDefaultExamRate.Text & "'," & _
                                    Val(chkWardActive.Value))
                                                    
                DeleteAVRate txtWardAvailableRate.Text, txtWardID.Text
                SaveAVRate txtWardAvailableRate.Text, txtWardID.Text
                
                tvwDeptWardsRooms.Nodes.Item(Val(lbltvwWardIndex.Caption)).Text = Trim(txtWardName.Text)
                tvwDeptWardsRooms.Nodes.Item(Val(lbltvwWardIndex.Caption)).Image = IIf(chkWardActive.Value = 1, "Ward", "Disabled")
        
                cmdWardClose_Click
            End If
        Else
            MsgBox "Please Fill Lab Code and Ward Name", vbInformation + vbOKOnly, "LMS"
        End If
    End If
End Sub

Private Sub cmdUpdateDept_Click()
    If Len(Trim(txtDeptName.Text)) > 0 Then
        If cmdUpdateDept.Caption = "Save" Then
            Set rsDeptWardsRooms = myConnection.Execute("EXEC sp_HospDepartment_Insert '" & txtDeptName.Text & "'," & _
                                   "'" & txtDeptDesc.Text & "',1")
            If rsDeptWardsRooms.Fields("InsertStat") = "2" Then
                MsgBox "Department " & txtDeptName.Text & " already exists"
            Else
                settvwDeptWardsRoomsData
            End If
            txtDeptID.Text = ""
            txtDeptName.Text = ""
            txtDeptDesc.Text = ""
            chkDeptActive.Value = 1
            txtDeptName.SetFocus
        Else
            Set rsDeptWardsRooms = myConnection.Execute("UPDATE tbHospDepartment SET Department = '" & txtDeptName.Text & "'," & _
                                    "Description = '" & txtDeptDesc.Text & "',Active = '" & chkDeptActive.Value & "'" & _
                                    " WHERE DeptID =" & txtDeptID.Text)
            tvwDeptWardsRooms.Nodes.Item(Val(lbltvwDeptIndex.Caption)).Text = Trim(txtDeptName.Text)
            tvwDeptWardsRooms.Nodes.Item(Val(lbltvwDeptIndex.Caption)).Image = IIf(chkDeptActive.Value = 1, "Close", "Disabled")
            
            cmdDeptClose_Click
        End If
    End If
End Sub

Private Sub cmdWardClose_Click()
    fraWards.Enabled = False
    txtDeptID.Text = ""
    txtDeptName.Text = ""
    txtDeptDesc.Text = ""
    chkDeptActive.Value = 0
    txtWardID.Text = ""
    txtWardLabCode.Text = ""
    txtWardHISCode.Text = ""
    txtWardName.Text = ""
    txtWardAvailableRate.Text = ""
    txtWardDefaultExamRate.Text = ""
    lbltvwWardIndex.Caption = ""
    chkWardActive.Value = 0
    cmdUpdateWard.Caption = "Save"
    cmdUpdateWard.Enabled = False
    cmdUpdateWard.Visible = False
    cmdWardClose.Enabled = False
    cmdWardClose.Visible = False
    tvwDeptWardsRooms.Enabled = True
    tvwDeptWardsRooms.SetFocus
End Sub

Private Sub Form_Load()
    Dim ctrImageList As Integer
    
    Set rsDeptWardsRooms = New ADODB.Recordset

    setFormColors Me
    fraDept.BackColor = GetSetting("LMS", "FormColors", "BodyFrame", &H6B4B01)
    fraDept.ForeColor = GetSetting("LMS", "FormColors", "TitleText", &HFFFFFF)
    fraWards.BackColor = GetSetting("LMS", "FormColors", "BodyFrame", &H6B4B01)
    fraWards.ForeColor = GetSetting("LMS", "FormColors", "TitleText", &HFFFFFF)
    fraRooms.BackColor = GetSetting("LMS", "FormColors", "BodyFrame", &H6B4B01)
    fraRooms.ForeColor = GetSetting("LMS", "FormColors", "TitleText", &HFFFFFF)
    
    mnuDept.Visible = False
    mnuWard.Visible = False
    mnuRoom.Visible = False

    Me.Height = 5775
    
    Set tvwDeptWardsRooms.ImageList = ImageList
    settvwDeptWardsRoomsData

End Sub

Private Sub mnuDeptAdd_Click()
    If tvwDeptWardsRooms.Nodes.Count = 0 Then
        txtDeptID.Text = ""
        txtDeptName.Text = ""
        chkDeptActive.Value = 0
        fraDept.Enabled = True
        fraWards.Enabled = False
        fraRooms.Enabled = False
        txtDeptName.Enabled = True
        txtDeptDesc.Enabled = True
        chkDeptActive.Enabled = True
        cmdUpdateDept.Caption = "Save"
        cmdUpdateDept.Visible = True
        cmdUpdateDept.Enabled = True
        cmdDeptClose.Visible = True
        cmdDeptClose.Enabled = True
        txtDeptName.SetFocus
        chkDeptActive.Value = 1
        tvwDeptWardsRooms.Enabled = False
    Else
        If tvwDeptWardsRooms.SelectedItem.Parent Is Nothing Then
            txtDeptID.Text = ""
            txtDeptName.Text = ""
            chkDeptActive.Value = 0
            fraDept.Enabled = True
            fraWards.Enabled = False
            fraRooms.Enabled = False
            txtDeptName.Enabled = True
            txtDeptDesc.Enabled = True
            chkDeptActive.Enabled = True
            cmdUpdateDept.Caption = "Save"
            cmdUpdateDept.Visible = True
            cmdUpdateDept.Enabled = True
            cmdDeptClose.Visible = True
            cmdDeptClose.Enabled = True
            txtDeptName.SetFocus
            chkDeptActive.Value = 1
            tvwDeptWardsRooms.Enabled = False
        End If
    End If
End Sub

Private Sub mnuDeptAddRoom_Click()
    If Left(tvwDeptWardsRooms.SelectedItem.Key, 1) = "D" Then
        tvwDeptWardsRooms.Enabled = False
        mnuDeptView_Click
        cmdDeptClose.Visible = False
        cmdDeptClose.Enabled = False
        fraDept.Enabled = False
        fraWards.Enabled = False
        fraRooms.Enabled = True
        txtRoomLabCode.SetFocus
        cmdUpdateRoom.Caption = "Save"
        cmdUpdateRoom.Visible = True
        cmdUpdateRoom.Enabled = True
        cmdRoomClose.Visible = True
        cmdRoomClose.Enabled = True
        chkRoomActive.Value = 1
    End If
End Sub

Private Sub mnuDeptEdit_Click()
    If tvwDeptWardsRooms.SelectedItem.Parent Is Nothing Then
        Set rsDeptWardsRooms = myConnection.Execute("SELECT * FROM tbHospDepartment WHERE DeptID = " & _
                                Mid(tvwDeptWardsRooms.SelectedItem.Key, 2))
        txtDeptID.Text = rsDeptWardsRooms.Fields("DeptID").Value
        txtDeptName.Text = rsDeptWardsRooms.Fields("Department").Value
        txtDeptDesc.Text = rsDeptWardsRooms.Fields("Description").Value
        chkDeptActive.Value = rsDeptWardsRooms.Fields("Active").Value
        tvwDeptWardsRooms.Enabled = False
        fraDept.Enabled = True
        fraWards.Enabled = False
        fraRooms.Enabled = False
        txtDeptName.Enabled = True
        chkDeptActive.Enabled = True
        txtDeptDesc.Enabled = True
        lbltvwDeptIndex.Caption = tvwDeptWardsRooms.SelectedItem.Index
        cmdUpdateDept.Caption = "Update"
        cmdUpdateDept.Visible = True
        cmdUpdateDept.Enabled = True
        cmdDeptClose.Visible = True
        cmdDeptClose.Enabled = True
    End If
End Sub

Private Sub mnuDeptView_Click()
    If tvwDeptWardsRooms.SelectedItem.Parent Is Nothing Then
        Set rsDeptWardsRooms = myConnection.Execute("SELECT * FROM tbHospDepartment WHERE DeptID = " & _
                                Mid(tvwDeptWardsRooms.SelectedItem.Key, 2))
        txtDeptID.Text = rsDeptWardsRooms.Fields("DeptID").Value
        txtDeptName.Text = rsDeptWardsRooms.Fields("Department").Value
        txtDeptDesc.Text = rsDeptWardsRooms.Fields("Description").Value
        chkDeptActive.Value = rsDeptWardsRooms.Fields("Active").Value
        tvwDeptWardsRooms.Enabled = False
        fraDept.Enabled = True
        fraWards.Enabled = False
        fraRooms.Enabled = False
        lbltvwDeptIndex.Caption = tvwDeptWardsRooms.SelectedItem.Index
        cmdDeptClose.Visible = True
        cmdDeptClose.Enabled = True
        cmdDeptClose.SetFocus
    End If
End Sub

Private Sub mnuEditRoom_Click()
    Dim ctrAVRate As Integer
    
    If Left(tvwDeptWardsRooms.SelectedItem.Key, 1) = "R" Then
        Set rsDeptWardsRooms = myConnection.Execute("SELECT a.RoomID,a.LISCode,a.HISCode,a.RoomNo," & _
                    "a.DefaultExamRate,a.Active FROM tbWardRooms AS a " & _
                    "WHERE a.Department = " & Mid(tvwDeptWardsRooms.SelectedItem.Parent.Key, 2) & _
                    " AND RoomID= " & Mid(tvwDeptWardsRooms.SelectedItem.Key, 2))
        txtRoomID.Text = rsDeptWardsRooms.Fields("RoomID").Value
        txtRoomLabCode.Text = rsDeptWardsRooms.Fields("LISCode").Value
        txtRoomHISCode.Text = rsDeptWardsRooms.Fields("HISCode").Value
        txtRoomNo.Text = rsDeptWardsRooms.Fields("RoomNo").Value
        chkRoomActive.Value = rsDeptWardsRooms.Fields("Active").Value
        txtRoomDefaultExamRate.Text = IIf(rsDeptWardsRooms.Fields("DefaultExamRate").Value = "0", "", rsDeptWardsRooms.Fields("DefaultExamRate").Value)
        Set rsDeptWardsRooms = myConnection.Execute("SELECT ExamRateID FROM tbAvailableRate WHERE RoomID =" & txtRoomID.Text)
        If rsDeptWardsRooms.RecordCount > 0 Then
            With rsDeptWardsRooms
                For ctrAVRate = 0 To .RecordCount - 1
                    txtRoomAvailableRate.Text = IIf(ctrAVRate = 0, .Fields("ExamRateID"), txtRoomAvailableRate.Text & _
                            "," & .Fields("ExamRateID"))
                    .MoveNext
                Next
            End With
        End If
        tvwDeptWardsRooms.Enabled = False
        fraDept.Enabled = False
        fraWards.Enabled = False
        fraRooms.Enabled = True
        lbltvwRoomIndex.Caption = tvwDeptWardsRooms.SelectedItem.Index
        cmdUpdateRoom.Caption = "Update"
        cmdUpdateRoom.Visible = True
        cmdUpdateRoom.Enabled = True
        cmdRoomClose.Visible = True
        cmdRoomClose.Enabled = True
        txtRoomLabCode.SetFocus
    End If
End Sub

Private Sub mnuRoomAdd_Click()
    If Left(tvwDeptWardsRooms.SelectedItem.Key, 1) = "W" Then
        tvwDeptWardsRooms.Enabled = False
        tvwDeptWardsRooms.SelectedItem.Expanded = False
        fraDept.Enabled = False
        fraWards.Enabled = False
        fraRooms.Enabled = True
        cmdUpdateRoom.Caption = "Save"
        cmdUpdateRoom.Enabled = True
        cmdUpdateRoom.Visible = True
        cmdRoomClose.Enabled = True
        cmdRoomClose.Visible = True
        chkRoomActive.Value = 1
    End If
End Sub

Private Sub mnuWardAdd_Click()
    If Left(tvwDeptWardsRooms.SelectedItem.Key, 1) = "D" Then
        tvwDeptWardsRooms.Enabled = False
        If Not tvwDeptWardsRooms.SelectedItem.Expanded Then
            tvwDeptWardsRooms.SelectedItem.Expanded = True
        End If
        mnuDeptView_Click
        cmdDeptClose.Visible = False
        cmdDeptClose.Enabled = False
        fraDept.Enabled = False
        fraWards.Enabled = True
        fraRooms.Enabled = False
        txtWardLabCode.SetFocus
        cmdUpdateWard.Caption = "Save"
        cmdUpdateWard.Visible = True
        cmdUpdateWard.Enabled = True
        cmdWardClose.Visible = True
        cmdWardClose.Enabled = True
        chkWardActive.Value = 1
    End If
End Sub

Private Sub mnuWardEdit_Click()
    Dim ctrAVRate As Integer
    
    If Left(tvwDeptWardsRooms.SelectedItem.Key, 1) = "W" Then
        Set rsDeptWardsRooms = myConnection.Execute("SELECT a.WardID,a.LISCode,a.HISCode,a.WardName" & _
                            ",a.DefaultExamRate,ISNULL(b.ExamClassName,'') AS ExamClassName,a.Active,a.SortOrder " & _
                            "FROM tbHospWards AS a LEFT OUTER JOIN tbLabExamRatesClass AS b ON a.DefaultExamRate = b.ExamClassID " & _
                            "WHERE a.Department = " & Mid(tvwDeptWardsRooms.SelectedItem.Parent.Key, 2) & _
                            " AND WardID = " & Mid(tvwDeptWardsRooms.SelectedItem.Key, 2))
        txtWardID.Text = rsDeptWardsRooms.Fields("WardID").Value
        txtWardLabCode.Text = rsDeptWardsRooms.Fields("LISCode").Value
        txtWardHISCode.Text = rsDeptWardsRooms.Fields("HISCode").Value
        txtWardName.Text = rsDeptWardsRooms.Fields("WardName").Value
        chkWardActive.Value = rsDeptWardsRooms.Fields("Active").Value
        txtWardDefaultExamRate.Text = IIf(rsDeptWardsRooms.Fields("DefaultExamRate").Value = "0", "", rsDeptWardsRooms.Fields("DefaultExamRate").Value)
        Set rsDeptWardsRooms = myConnection.Execute("SELECT ExamRateID FROM tbAvailableRate WHERE WardID =" & txtWardID.Text)
        If rsDeptWardsRooms.RecordCount > 0 Then
            With rsDeptWardsRooms
                For ctrAVRate = 0 To .RecordCount - 1
                    txtWardAvailableRate.Text = IIf(ctrAVRate = 0, .Fields("ExamRateID"), txtWardAvailableRate.Text & _
                            "," & .Fields("ExamRateID"))
                    .MoveNext
                Next
            End With
        End If
        tvwDeptWardsRooms.Enabled = False
        fraDept.Enabled = False
        fraWards.Enabled = True
        fraRooms.Enabled = False
        lbltvwWardIndex.Caption = tvwDeptWardsRooms.SelectedItem.Index
        cmdUpdateWard.Caption = "Update"
        cmdUpdateWard.Visible = True
        cmdUpdateWard.Enabled = True
        cmdWardClose.Visible = True
        cmdWardClose.Enabled = True
        txtWardLabCode.SetFocus
    End If
End Sub


Private Sub tvwDeptWardsRooms_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        If tvwDeptWardsRooms.Nodes.Count = 0 Then
            mnuDeptEdit.Visible = False
            mnuDeptView.Visible = False
            mnuWardAdd.Visible = False
            mnuDeptAddRoom.Visible = False
            PopupMenu mnuDept
            mnuDeptEdit.Visible = True
            mnuDeptView.Visible = True
            mnuWardAdd.Visible = True
            mnuDeptAddRoom.Visible = True
        Else
            If tvwDeptWardsRooms.SelectedItem.Parent Is Nothing Then
                PopupMenu mnuDept
            ElseIf Left(tvwDeptWardsRooms.SelectedItem.Key, 1) = "W" Then
                PopupMenu mnuWard
            ElseIf Left(tvwDeptWardsRooms.SelectedItem.Key, 1) = "R" Or Left(tvwDeptWardsRooms.SelectedItem.Key, 1) = "S" Then
                PopupMenu mnuRoom
            End If
        End If
    End If
End Sub

Private Sub txtDeptDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDeptName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtWardDept_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub settvwDeptWardsRoomsData()
    Dim rsDept As New ADODB.Recordset
    Dim rsWards As New ADODB.Recordset
    Dim rsRooms As New ADODB.Recordset
    
    Dim ctrDept As Integer
    Dim ctrWards As Integer
    Dim ctrRooms As Integer
    
    Set rsDept = myConnection.Execute("SELECT DeptID,Department,Active FROM tbHospDepartment")
    tvwDeptWardsRooms.Nodes.Clear
    If rsDept.RecordCount > 0 Then
        For ctrDept = 0 To rsDept.RecordCount - 1
            tvwDeptWardsRooms.Nodes.Add , , "D" & rsDept.Fields("DeptID"), rsDept.Fields("Department"), _
                    IIf(rsDept.Fields("Active") = "1", "Close", "Disabled")
            Set rsWards = myConnection.Execute("SELECT * FROM tbHospWards WHERE Department = " & rsDept.Fields("DeptID") & " ORDER BY WardName")
            If rsWards.RecordCount > 0 Then
                For ctrWards = 0 To rsWards.RecordCount - 1
                    tvwDeptWardsRooms.Nodes.Add "D" & rsDept.Fields("DeptID"), tvwChild, "W" & _
                                        rsWards.Fields("WardID"), rsWards.Fields("WardName"), _
                                        IIf(rsWards.Fields("Active") = "1", "Ward", "Disabled")
                    rsWards.MoveNext
                Next
            End If
            Set rsRooms = myConnection.Execute("SELECT RoomID,RoomNo,Active FROM tbWardRooms WHERE Department = " & rsDept.Fields("DeptID") & " ORDER BY RoomNo")
            If rsRooms.RecordCount > 0 Then
                For ctrRooms = 0 To rsRooms.RecordCount - 1
                    tvwDeptWardsRooms.Nodes.Add "D" & rsDept.Fields("DeptID"), tvwChild, "R" & _
                                    rsRooms.Fields("RoomID"), rsRooms.Fields("RoomNo"), _
                                    IIf(rsRooms.Fields("Active") = "1", "Room", "Disabled")
                    rsRooms.MoveNext
                Next
            End If
            rsDept.MoveNext
        Next
    End If

End Sub

Private Sub txtRoomAvailableRate_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strAvailableRate As String
    If KeyCode = 113 Then
        strAvailableRate = getExamRates(txtRoomAvailableRate.Text)
        If Len(strAvailableRate) > 0 Then
            txtRoomAvailableRate.Text = strAvailableRate
        End If
    End If
End Sub

Private Sub txtRoomAvailableRate_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRoomDefaultExamRate_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRoomHISCode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRoomLabCode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtWardAvailableRate_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strAvailableRate As String
    If KeyCode = 113 Then
        strAvailableRate = getExamRates(txtWardAvailableRate.Text)
        If Len(strAvailableRate) > 0 Then
            txtWardAvailableRate.Text = strAvailableRate
        End If
    End If
End Sub

Private Sub txtWardAvailableRate_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtWardDefaultExamRate_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtWardHISCode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtWardLabCode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtWardName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Function getExamRates(strPreRates As String) As String
    With frmLocAvailableRates
        .InitialRateID = strPreRates
        .Show 1
        getExamRates = .SelectedRates
    End With
End Function

Private Function checkDefRate(AVRate As String, DefRate As String) As Boolean
    Dim arryAVRate() As String
    Dim ctrAVRate As Integer
    arryAVRate = Split(AVRate, ",")
    
    checkDefRate = False
    For ctrAVRate = 0 To UBound(arryAVRate)
        If Trim(DefRate) = arryAVRate(ctrAVRate) Then
            checkDefRate = True
            Exit For
        End If
    Next
End Function

Private Sub SaveAVRate(AVRate As String, Optional WardID As String, Optional RoomID As String)
    Dim arryAVRate() As String
    Dim ctrAVRate As Integer
    Dim rsAVRate As New ADODB.Recordset
    Dim rsAVRateChecker As New ADODB.Recordset

    If Right(AVRate, "1") = "," Then
        AVRate = Mid(AVRate, 1, Len(AVRate) - 1)
    End If
    
    arryAVRate = Split(AVRate, ",")
    
    If Val(WardID) > 0 Then
        For ctrAVRate = 0 To UBound(arryAVRate)
            Set rsAVRate = myConnection.Execute("SELECT ExamClassID FROM tbLabExamRatesClass WHERE ExamClassID =" & arryAVRate(ctrAVRate))
            If rsAVRate.RecordCount > 0 Then
                Set rsAVRateChecker = myConnection.Execute("SELECT WardID FROM tbAvailableRate " & _
                            "WHERE WardID=" & WardID & " AND ExamRateID=" & arryAVRate(ctrAVRate))
                If rsAVRateChecker.RecordCount = 0 Then
                    myConnection.Execute ("INSERT INTO tbAvailableRate(WardID,RoomID,ExamRateID) VALUES" & _
                                "(" & WardID & ",0," & arryAVRate(ctrAVRate) & ")")
                End If
            End If
        Next
    ElseIf Val(RoomID) > 0 Then
        For ctrAVRate = 0 To UBound(arryAVRate)
            Set rsAVRate = myConnection.Execute("SELECT ExamClassID FROM tbLabExamRatesClass WHERE ExamClassID =" & arryAVRate(ctrAVRate))
            If rsAVRate.RecordCount > 0 Then
                Set rsAVRateChecker = myConnection.Execute("SELECT RoomID FROM tbAvailableRate " & _
                            "WHERE RoomID=" & RoomID & " AND ExamRateID=" & arryAVRate(ctrAVRate))
                If rsAVRateChecker.RecordCount = 0 Then
                    myConnection.Execute ("INSERT INTO tbAvailableRate(WardID,RoomID,ExamRateID) VALUES" & _
                                "(0," & RoomID & "," & arryAVRate(ctrAVRate) & ")")
                End If
            End If
        Next
    End If
    
End Sub

Private Sub DeleteAVRate(AVRate As String, Optional WardID As String, Optional RoomID As String)
    Dim arryAVRate() As String
    Dim ctrAVRate As Integer
    Dim rsAVRate As New ADODB.Recordset
    Dim strScript As String
    
    If Right(AVRate, 1) = "," Then
        AVRate = Mid(AVRate, 1, Len(AVRate) - 1)
    End If
    
    arryAVRate = Split(AVRate, ",")

    If Val(WardID) > 0 Then
        strScript = "DELETE FROM tbAvailableRate WHERE WardID=" & WardID & " AND ExamRateID NOT IN (" & AVRate & ")"
        myConnection.Execute (strScript)
    ElseIf Val(RoomID) > 0 Then
        strScript = "DELETE FROM tbAvailableRate WHERE RoomID=" & RoomID & " AND ExamRateID NOT IN (" & AVRate & ")"
        myConnection.Execute (strScript)
    End If
    
End Sub

