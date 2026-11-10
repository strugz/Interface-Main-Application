VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSelection 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selection"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.VScrollBar VScroll 
      Enabled         =   0   'False
      Height          =   3300
      Left            =   7305
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   150
      Width           =   255
   End
   Begin VB.Frame fraBody 
      Height          =   4125
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   -75
      Width           =   7665
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Height          =   405
         Left            =   5550
         TabIndex        =   1
         Top             =   3615
         Width           =   945
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   6615
         TabIndex        =   2
         Top             =   3615
         Width           =   945
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   3330
         Left            =   105
         TabIndex        =   0
         Top             =   210
         Width           =   7470
         _ExtentX        =   13176
         _ExtentY        =   5874
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         FocusRect       =   0
         ScrollBars      =   2
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
         _Band(0).Cols   =   1
      End
      Begin VB.Label lblText 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No Record Found..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   255
         TabIndex        =   4
         Top             =   3600
         Width           =   2925
      End
   End
End
Attribute VB_Name = "frmSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private returnCode As String
Private strFilter As String
Private strFilter2 As String
Private strFilter3 As String
Private strFilter4 As String
Private strFilter5 As String
Private strDataTrans As String
Private blnIsLoaded As Boolean
Private rsSearch As New ADODB.Recordset

Public Property Get setReturn() As String
    setReturn = returnCode
End Property

Public Property Let setDataTrans(ByVal vNewValue As String)
    strDataTrans = vNewValue
End Property

Public Property Let setFilter(ByVal vNewValue As String)
    strFilter = vNewValue
End Property

Public Property Let setFilter2(ByVal vNewValue As String)
    strFilter2 = vNewValue
End Property

Public Property Let setFilter3(ByVal vNewValue As String)
    strFilter3 = vNewValue
End Property

Public Property Let setFilter4(ByVal vNewValue As String)
    strFilter4 = vNewValue
End Property

Public Property Let setFilter5(ByVal vNewValue As String)
    strFilter5 = vNewValue
End Property

Public Property Let setRS(ByVal vNewValue As ADODB.Recordset)
    Set rsSearch.DataSource = vNewValue
End Property

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    With msfgResult
        If .Visible = True And cmdSelect.Enabled = True Then
            For intColCtr = 0 To .Cols - 1
                returnCode = returnCode & .TextMatrix(.Row, intColCtr) & "|"
            Next
            returnCode = Mid(returnCode, 1, Len(returnCode) - 1)
        End If
    End With
    Unload Me
End Sub

Private Sub Form_Load()

    blnIsLoaded = False
    
    setFormColors Me
    
    lblText(1).Move (Me.Width / 2) - (lblText(1).Width / 2), ((Me.Height / 2) - (lblText(1).Height / 2)) - 400

    returnCode = ""
    setGrid
    
    blnIsLoaded = True
    
End Sub

Private Sub setGrid()

    Dim strSQLString As String
    Dim intColCtr As Integer
      
    If InStr(1, strFilter, "'") > 0 Then
        strFilter = Replace(strFilter, "'", "''")
    End If
    
    If rsSearch.State = 0 Then
        strSQLString = "EXEC sp_SearchCodes2 " & strDataTrans & ",'" & strFilter & "','" & strFilter2 & "','" & _
                    strFilter3 & "','" & strFilter4 & "','" & strFilter5 & "'"
        Set rsSearch = myConnection.Execute(strSQLString)
    End If
    
    With msfgResult

        If rsSearch.RecordCount > 0 Then
            Set .DataSource = rsSearch
            If strDataTrans = 1 Then
                .ColWidth(0) = 1200
                .ColWidth(1) = 5980
            ElseIf strDataTrans = 2 Or strDataTrans = 3 Or strDataTrans = 4 Then
                .ColWidth(0) = 0
                .ColWidth(1) = 1200
                .ColWidth(2) = 2000
                .ColWidth(3) = 0
                .ColWidth(4) = 1980
                .ColWidth(5) = 2000
            ElseIf strDataTrans = 5 Or strDataTrans = 6 Then
                .ColWidth(0) = 0
                .ColWidth(1) = 7180
            End If
            
            For intColCtr = 0 To .Rows - 1
                .RowHeight(intColCtr) = 330
            Next
            
            If .Rows > 9 Then
                VScroll.Visible = False
            End If
            
            .Visible = True
            .Enabled = True
            
            cmdSelect.Enabled = True
            
            .ColSel = .Cols - 1
        Else
            .Visible = False
            .Enabled = False
            VScroll.Visible = False
            cmdSelect.Enabled = False
            cmdSelect.Visible = False
            cmdCancel.Caption = "Close"
        End If
    End With
    
    rsSearch.Close

End Sub


Private Sub Form_Unload(Cancel As Integer)
    If InStr(1, returnCode, "|") = 0 Then
        returnCode = "CANCELED"
    End If
End Sub

Private Sub msfgResult_DblClick()
    cmdSelect_Click
End Sub

Private Sub msfgResult_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSelect_Click
    End If
End Sub

Private Sub txtFilter_Change()
    If blnIsLoaded Then
        setGrid
    End If
End Sub

Private Sub txtFilter_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    With msfgResult
        Select Case KeyCode
            Case vbKeyDown
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    .ColSel = .Cols - 1
                Else
                    .Row = 1
                    .ColSel = .Cols - 1
                    .TopRow = .Row
                End If
                If Not .RowIsVisible(.Row) Then
                    .TopRow = .Row
                End If
                KeyCode = 0
                highLight
            Case vbKeyUp
                If .Row <> 1 Then
                    .Row = .Row - 1
                    .ColSel = .Cols - 1
                Else
                    .Row = .Rows - 1
                    .ColSel = .Cols - 1
                    .TopRow = .Row
                End If
                If Not .RowIsVisible(.Row) Then
                    .TopRow = .TopRow - 1
                End If
                KeyCode = 0
                highLight
            Case vbKeyPageDown
                If .Row + 12 <= .Rows Then
                    .Row = .Row + 12
                    .ColSel = .Cols - 1
                End If
                If Not .RowIsVisible(.Row) Then
                    .TopRow = .TopRow + 12
                End If
                KeyCode = 0
                highLight
            Case vbKeyPageUp
                If .Row <> 1 Then
                    If .TopRow < 12 Then
                        .TopRow = 1
                        .ColSel = .Cols - 1
                    Else
                        .Row = .Row - 12
                        If .Row < 12 Then
                            .TopRow = 1
                        End If
                        .ColSel = .Cols - 1
                    End If
                End If
                If Not .RowIsVisible(.Row) Then
                    If .TopRow > 12 Then
                        .TopRow = .TopRow - 12
                    End If
                End If
                KeyCode = 0
                highLight
        End Select
    End With
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSelect_Click
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub setListCmbSelection()
    Dim strFilter As String
    strFilter = GetSetting("HostForms_RE", "Settings\Others", "SpecList", "EXM_SPECIMEN")
    Set rsSearch = myConnection.Execute("exec sp_tbSelectionValues_select '" & strFilter & "'")
    If rsSearch.RecordCount > 0 Then
        For intRowCtr = 0 To rsSearch.RecordCount - 1
            cmbSpecimen.AddItem (rsSearch.Fields(1))
            rsSearch.MoveNext
        Next
    End If
End Sub
