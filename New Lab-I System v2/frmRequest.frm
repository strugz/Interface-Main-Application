VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRequest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Request Form"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13515
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   13515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame fraBody 
      BackColor       =   &H006B4B01&
      Height          =   7455
      Index           =   0
      Left            =   30
      TabIndex        =   19
      Top             =   -120
      Width           =   13515
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   345
         Left            =   6540
         TabIndex        =   14
         Top             =   270
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         _Version        =   393216
         Format          =   100007937
         CurrentDate     =   40441
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   405
         Left            =   8670
         TabIndex        =   15
         Top             =   240
         Width           =   1275
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   405
         Left            =   11970
         TabIndex        =   18
         Top             =   6780
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   4890
         TabIndex        =   13
         Top             =   270
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         _Version        =   393216
         Format          =   100007937
         CurrentDate     =   40437
      End
      Begin VB.Frame fraBody 
         BackColor       =   &H006B4B01&
         Height          =   6405
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   90
         Width           =   4065
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   405
            Left            =   2610
            TabIndex        =   12
            Top             =   4740
            Width           =   1275
         End
         Begin VB.TextBox txtRoom 
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
            Left            =   2880
            MaxLength       =   6
            TabIndex        =   11
            Top             =   4170
            Width           =   1095
         End
         Begin VB.TextBox txtWard 
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
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   10
            Top             =   4170
            Width           =   1095
         End
         Begin VB.TextBox txtSex 
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
            Left            =   2880
            MaxLength       =   6
            TabIndex        =   9
            Top             =   3780
            Width           =   1095
         End
         Begin VB.TextBox txtAge 
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
            Left            =   1170
            TabIndex        =   8
            Top             =   3780
            Width           =   1095
         End
         Begin VB.TextBox txtSpecimen 
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
            Left            =   1170
            MaxLength       =   15
            TabIndex        =   7
            Top             =   3390
            Width           =   2805
         End
         Begin VB.TextBox txtCivilStatus 
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
            Left            =   1170
            MaxLength       =   15
            TabIndex        =   6
            Top             =   3000
            Width           =   2805
         End
         Begin VB.TextBox txtBirthDate 
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
            Left            =   1170
            MaxLength       =   15
            TabIndex        =   5
            Top             =   2610
            Width           =   2805
         End
         Begin VB.TextBox txtMiddleName 
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
            Left            =   1170
            MaxLength       =   15
            TabIndex        =   4
            Top             =   2220
            Width           =   2805
         End
         Begin VB.TextBox txtFirstName 
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
            Left            =   1170
            MaxLength       =   15
            TabIndex        =   3
            Top             =   1830
            Width           =   2805
         End
         Begin VB.TextBox txtLastName 
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
            Left            =   1170
            MaxLength       =   15
            TabIndex        =   2
            Top             =   1440
            Width           =   2805
         End
         Begin VB.TextBox txtStatRoutine 
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
            Left            =   1170
            MaxLength       =   15
            TabIndex        =   1
            Top             =   1050
            Width           =   2805
         End
         Begin VB.TextBox txtSpecimenID 
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
            Left            =   1170
            MaxLength       =   15
            TabIndex        =   0
            Top             =   660
            Width           =   2805
         End
         Begin VB.TextBox txtPatientID 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   17
            Top             =   300
            Width           =   2805
         End
         Begin VB.Label lbltext 
            Alignment       =   1  'Right Justify
            BackColor       =   &H006B4B01&
            BackStyle       =   0  'Transparent
            Caption         =   "Age:"
            Height          =   195
            Index           =   8
            Left            =   90
            TabIndex        =   33
            Top             =   3840
            Width           =   975
         End
         Begin VB.Label lbltext 
            Alignment       =   1  'Right Justify
            BackColor       =   &H006B4B01&
            BackStyle       =   0  'Transparent
            Caption         =   "Civil Status:"
            Height          =   195
            Index           =   7
            Left            =   90
            TabIndex        =   32
            Top             =   3030
            Width           =   975
         End
         Begin VB.Label lbltext 
            Alignment       =   1  'Right Justify
            BackColor       =   &H006B4B01&
            BackStyle       =   0  'Transparent
            Caption         =   "Sex:"
            Height          =   195
            Index           =   9
            Left            =   1830
            TabIndex        =   31
            Top             =   3810
            Width           =   975
         End
         Begin VB.Label lbltext 
            Alignment       =   1  'Right Justify
            BackColor       =   &H006B4B01&
            BackStyle       =   0  'Transparent
            Caption         =   "Ward:"
            Height          =   195
            Index           =   10
            Left            =   90
            TabIndex        =   30
            Top             =   4230
            Width           =   975
         End
         Begin VB.Label lbltext 
            Alignment       =   1  'Right Justify
            BackColor       =   &H006B4B01&
            BackStyle       =   0  'Transparent
            Caption         =   "Room:"
            Height          =   195
            Index           =   11
            Left            =   1830
            TabIndex        =   29
            Top             =   4200
            Width           =   975
         End
         Begin VB.Label lbtext 
            Alignment       =   1  'Right Justify
            BackColor       =   &H006B4B01&
            BackStyle       =   0  'Transparent
            Caption         =   "Specimen:"
            Height          =   195
            Left            =   90
            TabIndex        =   28
            Top             =   3420
            Width           =   975
         End
         Begin VB.Label lbltext 
            Alignment       =   1  'Right Justify
            BackColor       =   &H006B4B01&
            BackStyle       =   0  'Transparent
            Caption         =   "Middle Name:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   27
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label lbltext 
            Alignment       =   1  'Right Justify
            BackColor       =   &H006B4B01&
            BackStyle       =   0  'Transparent
            Caption         =   "Last Name:"
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   26
            Top             =   1500
            Width           =   975
         End
         Begin VB.Label lbltext 
            Alignment       =   1  'Right Justify
            BackColor       =   &H006B4B01&
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Date:"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   25
            Top             =   2670
            Width           =   975
         End
         Begin VB.Label lbltext 
            Alignment       =   1  'Right Justify
            BackColor       =   &H006B4B01&
            BackStyle       =   0  'Transparent
            Caption         =   "First Name:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   24
            Top             =   1890
            Width           =   975
         End
         Begin VB.Label lbltext 
            Alignment       =   1  'Right Justify
            BackColor       =   &H006B4B01&
            BackStyle       =   0  'Transparent
            Caption         =   "Stat/Routine:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   1110
            Width           =   975
         End
         Begin VB.Label lbltext 
            Alignment       =   1  'Right Justify
            BackColor       =   &H006B4B01&
            BackStyle       =   0  'Transparent
            Caption         =   "Specimen ID:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   690
            Width           =   975
         End
         Begin VB.Label lbltext 
            Alignment       =   1  'Right Justify
            BackColor       =   &H006B4B01&
            BackStyle       =   0  'Transparent
            Caption         =   "Patient ID:"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   21
            Top             =   330
            Width           =   975
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgDisplay 
         Height          =   5745
         Left            =   4380
         TabIndex        =   16
         Top             =   810
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   10134
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   12615680
         ForeColorFixed  =   16777215
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lbltext 
         Alignment       =   1  'Right Justify
         BackColor       =   &H006B4B01&
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         Height          =   225
         Index           =   12
         Left            =   3840
         TabIndex        =   34
         Top             =   330
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstRequestForm As New ADODB.Recordset

Private Sub cmdAdd_Click()
If cmdAdd.Caption = "Add" Then
        If Len(txtLastName.Text) > 0 And Len(txtFirstName.Text) > 0 Then
            myConnection.Execute "exec sp_InsertRequestForm " & "'" & txtSpecimenID.Text & "','" & txtStatRoutine.Text & "', '" & txtLastName.Text & "', '" & txtFirstName.Text & "','" & txtMiddleName.Text & "','" & txtBirthDate.Text & "','" & txtCivilStatus.Text & "', '" & txtSpecimen.Text & "','" & txtAge.Text & "', '" & txtSex.Text & "', '" & txtWard & "','" & txtRoom.Text & "'"
        Else
            MsgBox "Invalid Entry"
        End If
      
Else
        myConnection.Execute "exec sp_UpdateRequestForm " & "'" & txtPatientID.Text & "','" & txtSpecimenID.Text & " ','" & txtStatRoutine.Text & "', '" & txtLastName.Text & "', '" & txtFirstName.Text & "','" & txtMiddleName.Text & "','" & txtBirthDate.Text & "','" & txtCivilStatus.Text & "', '" & txtSpecimen.Text & "','" & txtAge.Text & "', '" & txtSex.Text & "', '" & txtWard & "','" & txtRoom.Text & "'"
        cmdAdd.Caption = "Add"
        txtPatientID.Text = ""

End If
    Set msfgDisplay.Recordset = myConnection.Execute("Select * from tbLabRequestForm")
        txtSpecimenID.Text = ""
        txtStatRoutine.Text = ""
        txtLastName.Text = ""
        txtFirstName.Text = ""
        txtMiddleName.Text = ""
        txtBirthDate.Text = ""
        txtCivilStatus.Text = ""
        txtSpecimen.Text = ""
        txtAge.Text = ""
        txtSex.Text = ""
        txtWard.Text = ""
        txtRoom.Text = ""
        '" insert into  tbLabRequestForm (SpecimenID,StatRoutine,LastName,FirstName,MiddleName,BirthDate,Specimen,Age,Sex,Ward,Room) select '" & txtSpecimenID.Text & "','" & txtStatRoutine.Text & "','" & txtLastName.Text & "','" & txtFirstName.Text & "','" & txtMiddleName.Text & "', '" & txtBirthDate.Text & '",'" & txtCivilStatus.text & "',
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    setFormColors Me
    DTPicker1.Value = Now
    DTPicker2.Value = Now
    Set msfgDisplay.Recordset = myConnection.Execute("Select * from tbLabRequestForm")

End Sub
Private Sub msfgDisplay_DblClick()
    With msfgDisplay
            txtPatientID.Text = .TextMatrix(.Row, 0)
            txtSpecimenID.Text = .TextMatrix(.Row, 1)
            txtStatRoutine.Text = .TextMatrix(.Row, 2)
            txtLastName.Text = .TextMatrix(.Row, 3)
            txtFirstName.Text = .TextMatrix(.Row, 4)
            txtMiddleName.Text = .TextMatrix(.Row, 5)
            txtBirthDate.Text = .TextMatrix(.Row, 6)
            txtCivilStatus.Text = .TextMatrix(.Row, 7)
            txtSpecimen.Text = .TextMatrix(.Row, 8)
            txtAge.Text = .TextMatrix(.Row, 9)
            txtSex.Text = .TextMatrix(.Row, 10)
            txtWard.Text = .TextMatrix(.Row, 11)
            txtRoom.Text = .TextMatrix(.Row, 12)
            cmdAdd.Caption = "Update"
    End With
End Sub

Private Sub txtAge_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtBirthDate_Change()
Dim xIsDateMonths As String
Dim xYears As String
Dim xDays As String
Dim xMonths As String
'Dim tYears As Integer
'Dim tDays As Integer
'Dim tMonths As Integer
'Dim tempAge As Integer
'If Val(GetSetting("HostForms_SA", "Settings", "ComputeAge", "1")) = 1 Then
    If IsDate(txtBirthDate) Then
        xYears = DateDiff("yyyy", txtBirthDate.Text, Date)
        xMonths = DateDiff("m", txtBirthDate.Text, Date)
        xDays = DateDiff("d", txtBirthDate.Text, Date)
        'If xDays = 0 And xMonths = 0 And xYears = 0 Then
            txtAge.Text = "NEWBORN"
        'Else
            If xYears = 1 Then
                If Month(Now) = Month(txtBirthDate.Text) And Day(Now) >= Day(txtBirthDate.Text) Then
                    txtAge.Text = xYears
                ElseIf Month(Now) = Month(txtBirthDate.Text) And Day(Now) < Day(txtBirthDate.Text) Then
                    txtAge.Text = xMonths - 1 & IIf(xMonths = 1, " MONTH", " MONTHS")
                ElseIf Month(Now) < Month(txtBirthDate.Text) And Day(Now) >= Day(txtBirthDate.Text) Then
                    txtAge.Text = xMonths & IIf(xMonths = 1, " MONTH", " MONTHS")
                ElseIf Month(Now) < Month(txtBirthDate.Text) And Day(Now) < Day(txtBirthDate.Text) Then
                    txtAge.Text = (xMonths - 1) & IIf(xMonths = 1, " MONTH", " MONTHS")
                ElseIf Month(Now) > Month(txtBirthDate.Text) And Day(Now) >= Day(txtBirthDate.Text) Then
                    txtAge.Text = xYears
                ElseIf Month(Now) > Month(txtBirthDate.Text) And Day(Now) < Day(txtBirthDate.Text) Then
                    txtAge.Text = xYears
                End If
            ElseIf xYears > 1 Then
                If Month(Now) = Month(txtBirthDate.Text) And Day(Now) >= Day(txtBirthDate.Text) Then
                    txtAge.Text = xYears
                ElseIf Month(Now) = Month(txtBirthDate.Text) And Day(Now) < Day(txtBirthDate.Text) Then
                    txtAge.Text = xYears - 1
                ElseIf Month(Now) < Month(txtBirthDate.Text) And Day(Now) >= Day(txtBirthDate.Text) Then
                    txtAge.Text = xYears - 1
                ElseIf Month(Now) < Month(txtBirthDate.Text) And Day(Now) < Day(txtBirthDate.Text) Then
                    txtAge.Text = xYears - 1
                ElseIf Month(Now) > Month(txtBirthDate.Text) And Day(Now) >= Day(txtBirthDate.Text) Then
                    txtAge.Text = xYears
                ElseIf Month(Now) > Month(txtBirthDate.Text) And Day(Now) < Day(txtBirthDate.Text) Then
                    txtAge.Text = xYears
                End If
            ElseIf xYears < 1 Then
                If Month(Now) = Month(txtBirthDate.Text) And Day(Now) >= Day(txtBirthDate.Text) Then
                    If xDays < 1 Then
                        txtAge.Text = ""
                    Else
                        txtAge.Text = xDays & IIf(xDays = 1, " DAY", " DAYS")
                    End If
                ElseIf Month(Now) = Month(txtBirthDate.Text) And Day(Now) < Day(txtBirthDate.Text) Then
                    txtAge.Text = ""
                ElseIf Month(Now) < Month(txtBirthDate.Text) And Day(Now) >= Day(txtBirthDate.Text) Then
                    If xDays < 1 Then
                        txtAge.Text = ""
                    Else
                        txtAge.Text = xDays & IIf(xDays = 1, " DAY", " DAYS")
                    End If
                ElseIf Month(Now) < Month(txtBirthDate.Text) And Day(Now) < Day(txtBirthDate.Text) Then
                    If xDays < 1 Then
                        txtAge.Text = ""
                    Else
                        txtAge.Text = xDays & IIf(xDays = 1, " DAY", " DAYS")
                    End If
                ElseIf Month(Now) > Month(txtBirthDate.Text) And Day(Now) >= Day(txtBirthDate.Text) Then
                    If xMonths < 1 Then
                        txtAge.Text = ""
                    Else
                        txtAge.Text = xMonths & IIf(xMonths = 1, " MONTH", " MONTHS")
                    End If
                ElseIf Month(Now) > Month(txtBirthDate.Text) And Day(Now) < Day(txtBirthDate.Text) Then
                    If xDays < 1 Then
                        txtAge.Text = ""
                    Else
                        txtAge.Text = xMonths & IIf(xMonths = 1, " MONTH", " MONTHS")
                    End If
                    Else
                        txtAge.Text = ""
                End If
            End If
        End If
End Sub

Private Sub txtBirthDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtBirthDate_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCivilStatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtCivilStatus_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFirstName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtLastName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtMiddleName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtMiddleName_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPatientID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtPatientID_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRoom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtRoom_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRoutine_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtRoutine_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtSex_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtSex_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(1, "MF", Chr(KeyAscii)) > 0 Then
        If Chr(KeyAscii) = "M" Then txtSex.Text = "MALE"
        If Chr(KeyAscii) = "F" Then txtSex.Text = "FEMALE"
    End If
    KeyAscii = 0
End Sub

Private Sub txtSpecimen_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtSpecimen_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSpecimenID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtSpecimenID_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtStatRoutine_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtStatRoutine_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtWard_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtWard_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
