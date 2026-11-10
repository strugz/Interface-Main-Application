VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSummaryRequest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Summary"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frabody 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   6405
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   -90
      Width           =   7815
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   7830
         TabIndex        =   1
         Top             =   5820
         Width           =   1155
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Height          =   405
         Left            =   6600
         TabIndex        =   0
         Top             =   5820
         Width           =   1125
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   5415
         Left            =   90
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   9551
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   -2147483634
         GridColor       =   16777215
         GridColorFixed  =   16777215
         GridColorUnpopulated=   8421504
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
   End
End
Attribute VB_Name = "frmSummaryRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsSummary As New ADODB.Recordset
Private strRequestID As String
Private strSectionID As String
Private strSampleID As String
Private toUnloadOwner As Integer

Private Sub cmdCancel_Click()
    'toUnloadOwner = 0
    Unload Me
End Sub

Private Sub cmdOk_Click()
    toUnloadOwner = 1
    Unload Me
End Sub

Private Sub Form_Load()
    setFormColors Me
    setMsfgSummary
End Sub

Private Sub setMsfgSummary()
    Dim intRowCtr As Integer
    Dim strBarCode As String
    Dim strGroupBarCode As String
    Dim strGroupSectionName As String
    
    With msfgResult
        Set rsSummary = myConnection.Execute("exec sp_tbLabRequestValues_select '" & strRequestID & "'")
        If rsSummary.RecordCount > 0 Then
            
            .TextArray(0) = "Section"
            .TextArray(1) = "Sample Number - Specimen"
            .TextArray(2) = "Test Name"
                
            .ColWidth(0) = 2200
            .ColWidth(1) = 3100
            .ColWidth(2) = 2000
            
        
            If Len(Trim(strSectionID)) > 0 And Len(Trim(strSampleID)) > 0 Then
                rsSummary.Filter = "RVSectionID = '" & strSectionID & "' AND RVSampleID = '" & strSampleID & "'"
            ElseIf Len(Trim(strSectionID)) > 0 And Len(Trim(strSampleID)) = 0 Then
                rsSummary.Filter = "RVSectionID = '" & strSectionID & "'"
            ElseIf Len(Trim(strSectionID)) = 0 And Len(Trim(strSampleID)) > 0 Then
                rsSummary.Filter = "RVSampleID = '" & strSampleID & "'"
            End If
                
            rsSummary.Sort = "RVSectionID, Specimen"
            For intRowCtr = 0 To rsSummary.RecordCount - 1
                If strGroupSectionName <> rsSummary.Fields("SectionName").Value Then
                    .TextMatrix(.Rows - 1, 0) = rsSummary.Fields("SectionName").Value
                End If
                strGroupSectionName = rsSummary.Fields("SectionName").Value
                If strGroupBarCode <> rsSummary.Fields("RVSampleID").Value & " - " & rsSummary.Fields("Specimen").Value Then
                    .Rows = .Rows + 1
                    .Col = 1
                    .ColAlignment = 0
                    .TextMatrix(.Rows - 1, 1) = rsSummary.Fields("RVSampleID").Value & " - " & rsSummary.Fields("Specimen").Value
                    .Rows = .Rows + 1
                End If
                strGroupBarCode = rsSummary.Fields("RVSampleID").Value & " - " & rsSummary.Fields("Specimen").Value

                .TextMatrix(.Rows - 1, 2) = rsSummary.Fields("ExamRequestName").Value
                .Rows = .Rows + 1
                rsSummary.MoveNext
            Next
        Else
            .Visible = False
        End If
    End With
End Sub

Public Property Let SetSampleID(ByVal vNewValue As String)
    strSampleID = vNewValue
End Property

Public Property Let SetSectionID(ByVal vNewValue As String)
    strSectionID = vNewValue
End Property

Public Property Let SetRequestID(ByVal vNewValue As String)
    strRequestID = vNewValue
End Property

Public Property Get toUnload() As Integer
    toUnload = toUnloadOwner
End Property


