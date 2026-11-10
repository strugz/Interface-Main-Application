VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSampleIn 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Samples"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   495
      Left            =   5565
      TabIndex        =   0
      Top             =   3270
      Width           =   1215
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   4005
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   -75
      Width           =   6885
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgSamples 
         Height          =   3075
         Left            =   75
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   5424
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483647
         GridColorFixed  =   -2147483640
         GridColorUnpopulated=   16777215
         HighLight       =   2
         GridLinesFixed  =   1
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame fraTitle 
      Caption         =   "Frame1"
      Height          =   1860
      Left            =   7530
      TabIndex        =   2
      Top             =   1905
      Width           =   1095
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "lblText"
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   1095
         Width           =   1095
      End
      Begin VB.Label lblStatus 
         Caption         =   "Label1"
         Height          =   315
         Left            =   270
         TabIndex        =   4
         Top             =   660
         Width           =   765
      End
      Begin VB.Label lblTitle 
         Caption         =   "Label1"
         Height          =   300
         Left            =   270
         TabIndex        =   3
         Top             =   240
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmSampleIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PatientRequestID As String
Private RequestUserID As String
Private RequestDate As String

Public Property Let setPReqID(ByVal vNewValue As String)
    PatientRequestID = vNewValue
End Property

Public Property Let setReqUserID(ByVal vNewValue As String)
    RequestUserID = vNewValue
End Property

Public Property Let setRequestDate(ByVal vNewValue As String)
    RequestDate = vNewValue
End Property

Private Sub cmdDone_Click()
    Dim rsSpecimenIn As New ADODB.Recordset
    Dim intmsfgSamplesCtr As Integer
    With msfgSamples
        For intmsfgSamplesCtr = 1 To .Rows - 1
            If .TextMatrix(.Row, .Cols - 1) = "þ" Then
                strSQLStatement = "EXEC sp_CheckInBySmpl_update '" & PatientRequestID & "'," & _
                                "'" & .TextMatrix(intmsfgSamplesCtr, 2) & "'," & _
                                "'" & .TextMatrix(intmsfgSamplesCtr, 1) & "'," & _
                                "'" & RequestDate & "'," & _
                                "'" & RequestUserID & "'"
                Set rsSpecimenIn = myConnection.Execute(strSQLStatement)
            End If
        Next
    End With

    Unload Me
End Sub

Private Sub Form_Load()
    setFormColors Me
    msfgSamples.BackColorFixed = GetSetting("LMS", "FormColors", "TitleBar", &HA66B2B)
    msfgSamples.ForeColorFixed = GetSetting("LMS", "FormColors", "TitleText", &HFFFFFF)
    setmsfgSamples
End Sub

Private Sub setmsfgSamples()
    Dim rsSamples As New ADODB.Recordset
    Dim intRowSamples As Integer
    Dim intColSamples As Integer

    strSQLStatement = "EXEC sp_GetRequestSpecimens_select '" & PatientRequestID & "'"
        
    Set rsSamples = myConnection.Execute(strSQLStatement)
        
    With msfgSamples
        .Cols = rsSamples.Fields.Count + 1
        .Rows = rsSamples.RecordCount + 1
        
        For intColSamples = 0 To .Cols - 1
            .ColWidth(intColSamples) = 0
        Next

        .TextArray(0) = "Section"
        .TextArray(1) = "Specimen"
        .TextArray(.Cols - 1) = "With Sample"
    
        .RowHeight(0) = 350
        .ColWidth(0) = 2500
        .ColWidth(1) = 2500
        .ColWidth(.Cols - 1) = 1400
        
        For intRowSamples = 0 To rsSamples.RecordCount - 1
            .TextMatrix(intRowSamples + 1, 0) = rsSamples.Fields("SectionName").Value
            .TextMatrix(intRowSamples + 1, 1) = rsSamples.Fields("Specimen").Value
            .TextMatrix(intRowSamples + 1, 2) = rsSamples.Fields("RVSectionID").Value
            .TextMatrix(intRowSamples + 1, 3) = rsSamples.Fields("RVRequestID").Value
            .Row = intRowSamples + 1
            .Col = .Cols - 1
            .CellFontName = "Wingdings"
            .CellFontBold = True
            .CellFontSize = 12
            .CellAlignment = 4
            .TextMatrix(intRowSamples + 1, .Cols - 1) = "þ"
            rsSamples.MoveNext
        Next
    End With
    
End Sub

Private Sub msfgSamples_Click()
    With msfgSamples
        If .Col = .Cols - 1 Then
            If .TextMatrix(.Row, .Col) = "þ" Then
                .TextMatrix(.Row, .Col) = "o"
            Else
                .TextMatrix(.Row, .Col) = "þ"
            End If
        End If
    End With
End Sub

Private Sub msfgSamples_KeyPress(KeyAscii As Integer)
    With msfgSamples
        If .Col = .Cols - 1 Then
            If KeyAscii = 13 Then
                If .TextMatrix(.Row, .Col) = "þ" Then
                    .TextMatrix(.Row, .Col) = "o"
                Else
                    .TextMatrix(.Row, .Col) = "þ"
                End If
                KeyAscii = 0
            End If
        End If
    End With
End Sub
