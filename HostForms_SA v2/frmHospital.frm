VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmHospital 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hospital Information"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6030
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6330
      TabIndex        =   0
      Top             =   2340
      Width           =   5895
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Build File - [Hospital Information]"
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
         Width           =   3180
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   3705
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   30
      Width           =   5895
      Begin VB.TextBox txtDescription 
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
         Height          =   1545
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   2040
         Width           =   5505
      End
      Begin VB.TextBox txtAddress2 
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
         Height          =   345
         Left            =   1710
         TabIndex        =   8
         Top             =   1260
         Width           =   4005
      End
      Begin VB.TextBox txtAddress1 
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
         Height          =   345
         Left            =   1710
         TabIndex        =   6
         Top             =   810
         Width           =   4005
      End
      Begin VB.TextBox txtHospitalName 
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
         Height          =   345
         Left            =   1500
         TabIndex        =   4
         Top             =   360
         Width           =   4215
      End
      Begin MSComDlg.CommonDialog cdOpen 
         Left            =   60
         Top             =   90
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Hospital Description:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   1770
         Width           =   1695
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Address (Line 2):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1425
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Address (Line 1):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   870
         Width           =   1425
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Hospital Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   420
         Width           =   1215
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   60
      TabIndex        =   11
      Top             =   3750
      Width           =   5895
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   345
         Left            =   3240
         TabIndex        =   13
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   345
         Left            =   4560
         TabIndex        =   12
         Top             =   180
         Width           =   1275
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Hospital Logo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   30
      TabIndex        =   15
      Top             =   0
      Width           =   1170
   End
   Begin VB.Image imLogo 
      BorderStyle     =   1  'Fixed Single
      Height          =   1995
      Left            =   6120
      Picture         =   "frmHospital.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1995
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Hospital Logo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   30
      TabIndex        =   14
      Top             =   0
      Width           =   1170
   End
End
Attribute VB_Name = "frmHospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstream As ADODB.Stream
Private rstHospital As ADODB.Recordset

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
If myConnection.Execute("Select * from tbHospitalInfo where hospitalID = '1'").RecordCount > 0 Then
    myConnection.Execute "Update tbHospitalInfo set Hospitalname = '" & replaceStr(txtHospitalName.Text) & "', HospitalAddress1 = '" & replaceStr(txtAddress1.Text) & _
    "', HospitalAddress2 = '" & replaceStr(txtAddress2.Text) & "', HospitalDesc = '" & replaceStr(txtDescription.Text) & "' where HospitalID = '1'"
Else
    myConnection.Execute "Insert into tbHospitalInfo (Hospitalname, HospitalAddress1, HospitalAddress2, HospitalDesc,HospitalID,Active)" & _
            "Select '" & replaceStr(txtHospitalName.Text) & "','" & replaceStr(txtAddress1.Text) & "','" & replaceStr(txtAddress2.Text) & "','" & replaceStr(txtDescription.Text) & "','" & 1 & "'," & "True"
End If

If Len(Trim(cdOpen.FileName)) > 0 Then
    myConnection.Execute "Update tbHospitalInfo set HospitalLogo = " & mstream.Read & " where hospitalID = '1'"
End If
MsgBox "Record has been successfully saved!", vbOKOnly + vbInformation, "Build File"
End Sub

Private Sub Form_Load()
Dim blobPic As BlobFieldObject
setFormColors Me
    Set rstHospital = myConnection.Execute("select * from tbHospitalInfo where HospitalID = '1'")
    If rstHospital.RecordCount > 0 Then
        txtHospitalName.Text = rstHospital.Fields("HospitalName").Value & ""
        txtAddress1.Text = rstHospital.Fields("HospitalAddress1").Value & ""
        txtAddress2.Text = rstHospital.Fields("HospitalAddress2").Value & ""
        txtDescription.Text = rstHospital.Fields("HospitalDesc").Value & ""
'        If Not IsNull(rstHospital.Fields("HospitalLogo").Value) Then
'            blobPic.Field.Value = rstHospital.Fields("HospitalLogo").Value
'            imLogo.Picture = rstHospital.Fields("HospitalLogo").Value
'        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rstHospital = Nothing
End Sub

Private Sub imLogo_Click()
Set mstream = New ADODB.Stream
mstream.Type = adTypeBinary
mstream.Open

With cdOpen
    .DialogTitle = "Open Signature"
    .FileName = ""
    .Filter = "Image (*.bmp)|*.bmp"
    .ShowOpen
'    If Len(.FileName) <> 0 Then
'        mstream.LoadFromFile .FileName
'        imLogo.Picture = mstream.Read  'LoadPicture(.FileName)
'    End If
End With
End Sub

Private Function replaceStr(tempStr As String) As String
    replaceStr = Replace(tempStr, "'", "''")
End Function
