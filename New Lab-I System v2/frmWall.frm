VERSION 5.00
Begin VB.Form frmWall 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   5670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7785
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   645
   End
   Begin VB.Image imgHospitalLogo 
      Height          =   2250
      Left            =   330
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2940
   End
   Begin VB.Image Image1 
      Height          =   2595
      Left            =   270
      Stretch         =   -1  'True
      Top             =   360
      Width           =   4755
   End
   Begin VB.Image Logo1 
      Height          =   1395
      Left            =   5670
      Top             =   1830
      Width           =   1725
   End
End
Attribute VB_Name = "frmWall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Error Resume Next
    Me.Height = mdiMain.ScaleHeight '- mdiMain.tbrMain.Height + 100
    Me.Width = mdiMain.ScaleWidth - frmSearch.Width
    Me.Left = frmSearch.Width
    Me.Top = 1
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & getIP
    Form_Resize
    Image1.Picture = LoadResPicture("WALLIMAGE", vbResBitmap)
    'LoadPicture(App.Path & "\" & "lab-i.jpg")
    'wbWall.Navigate "file:///D:/OneDrive/Willie/INTERFACE%20VB6/New%20Lab-I%20System%20v2/Lab-i.html" 'App.Path & "\Lab-i.html"
End Sub

Private Sub Form_Resize()
    Image1.Width = Me.Width
    Image1.Height = Me.Height
    Image1.Top = 1
    Image1.Left = 1
    lblVersion.Left = (Me.Width - lblVersion.Width) - 200
'    wbWall.Width = Me.Width
'    wbWall.Height = Me.Height
'    wbWall.Top = 1
'    wbWall.Left = 1
End Sub

Public Sub changeSize()
    On Error Resume Next
    Me.Height = mdiMain.ScaleHeight
    Me.Width = mdiMain.ScaleWidth - frmSearch.Width
    Me.Left = frmSearch.Width
    Me.Top = 1
    Form_Resize
End Sub

Private Function getIP(Optional getIPonly As Boolean = False) As String
Dim tempstr As String
Dim wmi     As Object
Dim qryWMI  As Object
Dim Item    As Variant
    tempstr = ""
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")

    Set qryWMI = wmi.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration " & _
                               "WHERE IPEnabled = True")

    For Each Item In qryWMI
        If Len(tempstr) = 0 Then
            If getIPonly Then
                tempstr = Item.IPAddress(0)
                Exit For
            Else
                tempstr = Item.IPAddress(0) & vbCrLf
            End If
        Else
            tempstr = tempstr & Item.IPAddress(0) & vbCrLf
        End If
    Next
    If Not getIPonly Then
        If Len(Trim(tempstr)) = 0 Then tempstr = "No IP detected or disconected from the network."
    End If
    Set wmi = Nothing
    Set qryWMI = Nothing
    getIP = tempstr
End Function

