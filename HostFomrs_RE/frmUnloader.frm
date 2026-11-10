VERSION 5.00
Begin VB.Form frmUnloader 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmUnloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    CloseForms
    Me.Visible = False
End Sub

Private Sub CloseForms()
   Dim ctr As Integer
    Dim tempStr As String
    Dim mhWnd() As Form
    
    searchObj.CloseForms
    
    ctr = 0
    tempStr = ""
    
    Dim frm As Form
    For Each frm In Forms
        If frm.hWnd <> Me.hWnd Then
             ctr = ctr + 1
        End If
    Next frm
    ReDim mhWnd(ctr) As Form
    ctr = 0
    For Each frm In Forms
        If frm.hWnd <> Me.hWnd Then
            ctr = ctr + 1
            Set mhWnd(ctr) = frm
        End If
    Next frm
    Do While ctr <> 0
         Unload mhWnd(ctr)
         Set mhWnd(ctr) = Nothing
         ctr = ctr - 1
    Loop
    
    isExiting = True
End Sub
