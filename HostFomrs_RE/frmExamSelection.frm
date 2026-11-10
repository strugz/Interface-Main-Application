VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmExamSelection 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Exam"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraBody 
      Height          =   4800
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   -75
      Width           =   7875
      Begin VB.ComboBox cmbSpecimen 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmExamSelection.frx":0000
         Left            =   5655
         List            =   "frmExamSelection.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   300
         Width           =   2145
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Height          =   405
         Left            =   5760
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   4275
         Width           =   945
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   6825
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   4275
         Width           =   945
      End
      Begin VB.TextBox txtFilter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   540
         MaxLength       =   30
         TabIndex        =   0
         Top             =   300
         Width           =   4155
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   3330
         Left            =   105
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   870
         Width           =   7680
         _ExtentX        =   13547
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
         BackStyle       =   0  'Transparent
         Caption         =   "Specimen"
         Height          =   285
         Index           =   2
         Left            =   4830
         TabIndex        =   7
         Top             =   360
         Width           =   735
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
         Left            =   1080
         TabIndex        =   6
         Top             =   1650
         Width           =   2925
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   30
         X2              =   7830
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Filter"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1365
      End
   End
End
Attribute VB_Name = "frmExamSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private returnCode As String
Private strFilter As String
Private strUserID As String
Private strUserType As String
Private strSpecimen As String
Private strSectionID As String
Private blnIsLoaded As Boolean
Private rsSearch As New ADODB.Recordset

Public Property Get setReturn() As String
    setReturn = returnCode
End Property

Public Property Let setFilter(ByVal vNewValue As String)
    strFilter = vNewValue
End Property

Public Property Let setUserID(ByVal vNewValue As String)
    strUserID = vNewValue
End Property

Public Property Let setUseType(ByVal vNewValue As String)
    strUserType = vNewValue
End Property

Public Property Let setSpecimen(ByVal vNewValue As String)
    strSpecimen = Trim(vNewValue)
End Property

Public Property Let setSectionID(ByVal vNewValue As String)
    strSectionID = Trim(vNewValue)
End Property

Private Sub cmbSpecimen_Click()
    If blnIsLoaded Then
        setGrid
        txtFilter.SetFocus
    End If
End Sub

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
    lblText(1).Move (Me.Width / 2) - (lblText(1).Width / 2), (Me.Height / 2) - (lblText(1).Height / 2)
    setListCmbSelection
    returnCode = ""

    If Len(Trim(strFilter)) > 0 Then
        txtFilter.Text = strFilter
    End If
    
    If Len(strSpecimen) > 0 Then
        For intColCtr = 0 To cmbSpecimen.ListCount - 1
            If Trim(cmbSpecimen.List(intColCtr)) = Trim(strSpecimen) Then
                cmbSpecimen.ListIndex = intColCtr
            End If
        Next
        cmbSpecimen.Enabled = False
    Else
        cmbSpecimen.ListIndex = 0
    End If
    
    setGrid
    
    blnIsLoaded = True
    
End Sub

Private Sub setGrid()

    Dim strSQLString As String
    Dim intColCtr As Integer
    
    strFilter = txtFilter.Text
    
    If InStr(1, strFilter, "'") > 0 Then
        strFilter = Replace(strFilter, "'", "''")
    End If
    
    strSQLString = "EXEC sp_ReqExam_Select " & strUserID & ",'" & strFilter & "'"
    
    Set rsSearch = myConnection.Execute(strSQLString)
    
    If cmbSpecimen.Text <> "All" Then
        If Len(strSectionID) > 0 Then
            rsSearch.Filter = "Specimen = '" & cmbSpecimen.Text & "' AND SectionID ='" & strSectionID & "'"
        Else
            rsSearch.Filter = "Specimen = '" & cmbSpecimen.Text & "'"
        End If
    End If

    txtFilter.SelStart = Len(txtFilter.Text)
    With msfgResult

        If rsSearch.RecordCount > 0 Then
            Set .DataSource = rsSearch
            .Visible = True
            .Enabled = True
            For intColCtr = 2 To .Cols - 1
                .ColWidth(intColCtr) = 0
            Next

            .ColWidth(0) = 600
            .ColWidth(1) = 3600
            .ColWidth(2) = 1500
            .ColWidth(5) = 1700
            .TextArray(0) = "Code"
            .TextArray(1) = "Name"
            .TextArray(2) = "Specimen"
            .TextArray(5) = "Section"
    
            cmdSelect.Enabled = True
            
            .ColSel = .Cols - 1
        Else
            .Visible = False
            .Enabled = False
            cmdSelect.Enabled = False
        End If
    End With
    
    rsSearch.Close

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
