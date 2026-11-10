VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmReportFormats 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Formats"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   13560
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   5430
      Width           =   2055
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Preview"
         Height          =   405
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1365
      End
      Begin VB.TextBox txtRptID 
         Alignment       =   1  'Right Justify
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
         Left            =   990
         MaxLength       =   8
         TabIndex        =   7
         Top             =   540
         Width           =   465
      End
      Begin RichTextLib.RichTextBox txtDesciption 
         Height          =   1215
         Left            =   990
         TabIndex        =   6
         Top             =   900
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   2143
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmReportFormats.frx":0000
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FCE2C7&
         BackStyle       =   0  'Transparent
         Caption         =   "Profile Exam:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   60
         TabIndex        =   10
         Top             =   570
         Width           =   930
      End
      Begin VB.Label lblRptName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1470
         TabIndex        =   9
         Top             =   540
         Width           =   3135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FCE2C7&
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   900
         Width           =   855
      End
   End
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   2430
      TabIndex        =   0
      Top             =   5460
      Width           =   3315
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Build File - [Report Formats]"
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
         Top             =   180
         Width           =   2715
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   4065
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   13425
      Begin VB.CommandButton cmdSave 
         Caption         =   "Set as default..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         TabIndex        =   39
         Top             =   3600
         Width           =   5295
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "Upload"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9990
         TabIndex        =   32
         Top             =   1410
         Width           =   1245
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8760
         TabIndex        =   31
         Top             =   1410
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2385
         Left            =   11310
         TabIndex        =   30
         Top             =   1410
         Width           =   945
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Save Report"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2385
         Left            =   12300
         TabIndex        =   29
         Top             =   1410
         Width           =   945
      End
      Begin VB.TextBox txtDisplayName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   6540
         TabIndex        =   24
         Top             =   1890
         Width           =   4695
      End
      Begin VB.TextBox txtReportName 
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
         Height          =   350
         Left            =   6540
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1020
         Width           =   4365
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
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
         Left            =   12090
         TabIndex        =   19
         Top             =   1020
         Width           =   1155
      End
      Begin VB.TextBox txtExtName 
         BackColor       =   &H8000000F&
         Height          =   345
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1020
         Width           =   1155
      End
      Begin VB.TextBox txtEditor 
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
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   210
         Width           =   6345
      End
      Begin VB.CommandButton cmdOpenEditor 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12870
         TabIndex        =   15
         Top             =   210
         Width           =   405
      End
      Begin MSComDlg.CommonDialog cdOpen 
         Left            =   12780
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   3405
         Left            =   90
         TabIndex        =   17
         Top             =   180
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   6006
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   12615680
         ForeColorFixed  =   16777215
         BackColorBkg    =   12632256
         GridColor       =   -2147483632
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
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
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin RichTextLib.RichTextBox txtNote 
         Height          =   615
         Left            =   6540
         TabIndex        =   26
         Top             =   2310
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   1085
         _Version        =   393217
         Enabled         =   -1  'True
         MaxLength       =   254
         Appearance      =   0
         TextRTF         =   $"frmReportFormats.frx":0082
      End
      Begin MSComDlg.CommonDialog cdBrowse 
         Left            =   12900
         Top             =   1380
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblModBy 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6540
         TabIndex        =   38
         Top             =   3420
         Width           =   4665
      End
      Begin VB.Label lblModDate 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6540
         TabIndex        =   37
         Top             =   3000
         Width           =   4665
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modified By:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   8
         Left            =   5610
         TabIndex        =   36
         Top             =   3480
         Width           =   885
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Modified:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   7
         Left            =   5490
         TabIndex        =   35
         Top             =   3090
         Width           =   1005
      End
      Begin VB.Label lblSize 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0K"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   6540
         TabIndex        =   34
         Top             =   1440
         Width           =   2205
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Size:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   4
         Left            =   5820
         TabIndex        =   33
         Top             =   1500
         Width           =   645
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NO RECORD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   180
         TabIndex        =   28
         Top             =   1830
         Width           =   5145
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Descrption:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   6
         Left            =   5640
         TabIndex        =   27
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Report Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   5
         Left            =   5490
         TabIndex        =   25
         Top             =   1950
         Width           =   1155
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Report File:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   3
         Left            =   5640
         TabIndex        =   23
         Top             =   1050
         Width           =   1005
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Report ID:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   2
         Left            =   5760
         TabIndex        =   22
         Top             =   690
         Width           =   765
      End
      Begin VB.Label lblReportID 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6540
         TabIndex        =   21
         Top             =   600
         Width           =   6675
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Editor:"
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
         Left            =   5910
         TabIndex        =   14
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lblReportFormat 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1440
         TabIndex        =   13
         Top             =   270
         Visible         =   0   'False
         Width           =   3915
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Report:"
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
         Left            =   180
         TabIndex        =   12
         Top             =   270
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   675
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   4170
      Width           =   13425
      Begin VB.CommandButton cmdOpenFolder 
         Cancel          =   -1  'True
         Caption         =   "Open Reports Folder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         TabIndex        =   41
         Top             =   180
         Width           =   2085
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11970
         TabIndex        =   4
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Before saving, click edit. Save and close the report. Click upload then save report."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2340
         TabIndex        =   40
         Top             =   210
         Visible         =   0   'False
         Width           =   9480
      End
   End
End
Attribute VB_Name = "frmReportFormats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FileStream As New ADODB.Stream

Private Sub cmdBrowse_Click()
Dim tempFilename As String
    cdBrowse.Filter = "RPT|" & SectionCode & "_*.rpt"
    If UCase(Dir(App.Path & "\Reports", vbDirectory)) = "REPORTS" Then
        cdBrowse.InitDir = App.Path + "\Reports"
    ElseIf UCase(Dir(Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports", vbDirectory)) = "REPORTS" Then
        cdBrowse.InitDir = Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports"
    End If
    cdBrowse.ShowOpen
    tempFilename = cdBrowse.Filename
    If Format(Right(tempFilename, 4), ">") = ".RPT" Then
        txtExtName.Text = "RPT"
    ElseIf Format(Right(tempFilename, 4), ">") = ".EXE" Then
        txtExtName.Text = "EXE"
    Else
        txtExtName.Text = ""
    End If
'    If UCase(txtReportName.Text) <> UCase(tempFilename) And Len(Trim(txtReportName.Text)) > 0 Then
'        if msgbox("Do you want to ")
'    End If
    
    lblSize.Caption = LoadReport2Stream(tempFilename)
    If lblSize.Caption <> "0K" Then
        cmdUpload.Enabled = False
        txtDisplayName.Text = Dir(tempFilename)
        cmdEdit.Enabled = True
        cmdCancel.Enabled = True
    Else
    End If
End Sub

Private Sub cmdCancel_Click()
    ClearValues
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
Select Case cmdEdit.Caption
    Case "Edit"
        If Len(Trim(txtDisplayName.Text)) > 0 Then
            FileStream.SaveToFile IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\" & txtDisplayName.Text, adSaveCreateOverWrite
            cmdUpload.Enabled = True
            Shell txtEditor.Text & " """ & IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\" & txtDisplayName.Text, vbMaximizedFocus
        Else
            FileStream.SaveToFile IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\" & "TEMP", adSaveCreateOverWrite
            cmdUpload.Enabled = True
            Shell txtEditor.Text & " """ & IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\TEMP", vbMaximizedFocus
        End If
    Case "Open"
        If Len(Trim(txtDisplayName.Text)) > 0 Then
            If Format(Right(txtDisplayName.Text, 4), ">") = "EXE" Then
                FileStream.SaveToFile IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\" & txtDisplayName.Text, adSaveCreateOverWrite
                cmdUpload.Enabled = True
                Shell IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\" & txtDisplayName.Text, vbMaximizedFocus
            Else
                FileStream.SaveToFile IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\" & txtDisplayName.Text & ".exe", adSaveCreateOverWrite
                cmdUpload.Enabled = True
                Shell IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\" & txtDisplayName.Text & ".EXE", vbNormalFocus
            End If
        Else
            FileStream.SaveToFile IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\" & "TEMP.EXE", adSaveCreateOverWrite
            cmdUpload.Enabled = True
            Shell IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\TEMP.EXE", vbMaximizedFocus
        End If
End Select

End Sub

Private Sub cmdOpenEditor_Click()
With cdOpen
    .DialogTitle = "Open Editor"
    If txtEditor.Text = "EDITOR NOT FOUND..." Then
        .InitDir = Environ("ProgramFiles")
    Else
        .InitDir = Left(txtEditor.Text, Len(txtEditor.Text) - 10)
    End If
    .Filter = "crw32.exe|crw32.exe"
    .ShowOpen
    If Len(cdOpen.Filename) > 0 Then
        txtEditor.Text = cdOpen.Filename
        SaveSetting "HostForms_SA", "REPORTEDITOR", "FILENAME", txtEditor.Text
    End If
End With
End Sub

Private Sub cmdOpenFolder_Click()
If UCase(Dir(App.Path & "\Reports", vbDirectory)) = "REPORTS" Then
    Shell "explorer " + App.Path & "\Reports", vbNormalFocus
ElseIf UCase(Dir(Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports", vbDirectory)) = "REPORTS" Then
    Shell "explorer " + Environ$("AllUsersProfile") & "\Laboratory Instrument Interface\Reports", vbNormalFocus
End If
End Sub

Private Sub cmdSave_Click()
Dim cRow As Integer
With msfgResult
    cRow = .Row
    myConnection.Execute "Update tbLabReportNames set rptDefault = False"
    myConnection.Execute "Update tbLabReportNames set rptDefault = True where rptID = " & .TextMatrix(.Row, 1)
    LoadGrid
    .Row = cRow
End With
End Sub

Private Sub fileList_DblClick()
'    If Len(Trim(fileList.FileName)) > 0 Or fileList.FileName <> "None" Then
'        OpenMainReport App.Path & "\reports\" & fileList.FileName, SectionCode & "_TESTPREVIEW"
'        ShowReportViewer False, "Result Print Preview", False
'    Else
'        MsgBox "Error"
'    End If
End Sub

Private Sub cmdUpdate_Click()
Dim mSaved As Boolean
mSaved = False
If cmdUpload.Enabled Then
    If MsgBox("Do you wish to upload and save modified report?", vbQuestion + vbYesNo, "Upload Report") = vbYes Then
        cmdUpload_Click
    Else
        MsgBox "You choose not not upload the edited report but saving other fields will continue.", vbOKOnly + vbInformation, "Reports"
    End If
End If
If Trim(lblReportID.Caption) = "" Then
    mSaved = SaveReport("I", txtReportName.Text, txtDisplayName.Text, txtNote.Text, SectionCode, , , True)
Else
    With msfgResult
        mSaved = SaveReport("U", txtReportName.Text, txtDisplayName.Text, txtNote.Text, SectionCode, , .TextMatrix(.Row, 1), True)
    End With
End If
If mSaved Then
    LoadGrid
    ClearValues
End If
End Sub

Private Sub cmdUpload_Click()
    If Len(Trim(txtReportName.Text)) > 0 Then
        If Len(Trim(Dir(IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\" & txtReportName.Text))) > 0 Then
            lblSize.Caption = LoadReport2Stream(IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\" & txtReportName.Text)
            Kill IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\" & txtReportName.Text
            cmdUpload.Enabled = False
        End If
    End If
End Sub

Private Sub Form_Load()
setFormColors Me
    Dim tempReport As String
    If Len(Dir(Environ("ProgramFiles") & "\Seagate Software\Crystal Reports\CRW32.exe")) > 0 Then
        txtEditor.Text = GetSetting("HostForms_SA", "REPORTEDITOR", "FILENAME", Environ("ProgramFiles") & "\Seagate Software\Crystal Reports\CRW32.exe")
        SaveSetting "HostForms_SA", "REPORTEDITOR", "FILENAME", txtEditor.Text
    Else
        If GetSetting("HostForms_SA", "REPORTEDITOR", "FILENAME", "") = "" Then
            txtEditor.Text = "EDITOR NOT FOUND..."
        Else
            txtEditor.Text = GetSetting("HostForms_SA", "REPORTEDITOR", "FILENAME", "")
        End If
    End If
    tempReport = GetSetting("HostForms_SA", "Settings", "ReportForm", "")
    If Len(Trim(tempReport)) > 0 Then
        lblReportFormat.Caption = tempReport
    Else
        lblReportFormat.Caption = "None"
    End If
    LoadGrid
    ClearValues
End Sub

Private Sub LoadGrid()
Dim rstTemp As ADODB.Recordset
Set rstTemp = myConnection.Execute("Select * from tbLabReportNames where rptSectionCode = '" & SectionCode & "'")
If rstTemp.RecordCount > 0 Then
    msfgResult.Visible = True
    Set msfgResult.DataSource = rstTemp
    With msfgResult
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = .Width - 300
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .TextMatrix(0, 2) = "Report Name"
        .Row = 1
        SetGridDefault .Row, True
    End With
Else
    msfgResult.Visible = False
    ImportReports
    Set rstTemp = myConnection.Execute("Select * from tbLabReportNames where rptSectionCode = '" & SectionCode & "'")
    If rstTemp.RecordCount > 0 Then
        msfgResult.Visible = True
        Set msfgResult.DataSource = rstTemp
        With msfgResult
            .ColWidth(0) = 0
            .ColWidth(1) = 0
            .ColWidth(2) = .Width - 300
            .ColWidth(3) = 0
            .ColWidth(4) = 0
            .ColWidth(5) = 0
            .ColWidth(6) = 0
            .ColWidth(7) = 0
            .ColWidth(8) = 0
            .TextMatrix(0, 2) = "Report Name"
            .Row = 1
            SetGridDefault .Row, True
        End With
    End If
End If
Set rstTemp = Nothing
End Sub

Private Function LoadReport2Stream(mFileLocation As String) As String
If FileStream.State = adStateOpen Then FileStream.Close
With FileStream
    .Type = adTypeBinary
    .Open
    .Flush
    If Len(Trim(Dir(mFileLocation))) > 0 And Len(mFileLocation) > 0 Then
        .LoadFromFile mFileLocation
        LoadReport2Stream = CStr(CInt((.Size / 1024))) & "KB"
    Else
        LoadReport2Stream = "0K"
    End If
End With
End Function

Private Sub lblSize_Change()
    If lblSize.Caption = "0K" Then
        cmdEdit.Enabled = False
    Else
        cmdEdit.Enabled = True
    End If
End Sub

Private Sub msfgResult_DblClick()
    msfgResult.Enabled = False
    cmdSave.Enabled = False
    cmdCancel.Cancel = True
    assignValues
End Sub

Private Sub txtDisplayName_Change()
    txtReportName.Text = txtDisplayName.Text
    If Len(Trim(txtDisplayName.Text)) > 0 Then
        cmdUpdate.Enabled = True
        cmdCancel.Enabled = True
    End If
End Sub

Private Sub ImportReports()
Dim rstTemp As ADODB.Recordset
Dim mReports  As String
Dim mReportsPath  As String
Dim tempStr As String
Dim ctr As Integer
ctr = 0
mReports = ""
mReportsPath = ""
tempStr = ""
Set rstTemp = myConnection.Execute("Select distinct template from tblabprofiles where instrumentid = '" & InstrumentID & "'")
If rstTemp.RecordCount > 0 Then
    If MsgBox("Do you want to import reports to database?", vbQuestion + vbYesNo, "Import Reports") = vbYes Then
        ReDim arrReports(rstTemp.RecordCount) As String
        rstTemp.MoveFirst
        Do While Not rstTemp.EOF
            If Len(Trim(rstTemp.Fields("Template").Value & "")) > 0 Then
                tempStr = Dir(App.Path & "\Reports\" & Trim(rstTemp.Fields("Template").Value & ""))
                mReportsPath = App.Path & "\Reports\" & Trim(rstTemp.Fields("Template").Value & "")
                mReports = tempStr
                SaveReport "I", mReportsPath, mReports, mReports, SectionCode, "AUTOUPLOAD"
            End If
            ctr = ctr + 1
            rstTemp.MoveNext
        Loop
    End If
End If
myConnection.Execute "Update tbLabReportNames set rptDefault = True where rptName = '" & GetSetting("HostForms_SA", "Settings", "ReportForm", "") & "'"
Set rstTemp = Nothing
End Sub

Private Function SaveReport(mTrans As String, rFname As String, rName As String, rDesc As String, rScode As String, Optional rUser As String = "", Optional rID As String = "", Optional PromptMessage As Boolean = False) As Boolean
Dim mycommand As New ADODB.Command
Dim mReportEnabled As String
Dim mUserInit As String
Dim mClass As New ClassForms
Dim Saved As Boolean
Dim myFile() As Byte
Saved = False
If mTrans = "I" And HasReport(rName) Then
    If PromptMessage Then MsgBox "Report Already Exists.", vbInformation + vbOKOnly, "Reports"
    Exit Function
End If
If mTrans = "I" Then
    If rName = "" Then
        If PromptMessage Then MsgBox "No report was selected", vbInformation + vbOKOnly, "Reports"
    Else
        If Len(Trim(rUser)) = 0 Then
            mUserInit = mClass.showInputBox("Input your initial then press Enter.", , , , 10, True, False, False)
        Else
            mUserInit = rUser
        End If
        If Len(mUserInit) > 0 And LoadReport2Stream(rFname) <> "0K" Then
            myFile = FileStream.Read(FileStream.Size)
            With mycommand
                .ActiveConnection = myConnection
                .CommandType = adCmdStoredProc
                .CommandText = "sp_ReportInsert"
                .Parameters.Append .CreateParameter("mRptName", adVarChar, adParamInput, 255)
                .Parameters.Append .CreateParameter("mRptDesc", adVarChar, adParamInput, 255)
                .Parameters.Append .CreateParameter("mRptSectionCode", adVarChar, adParamInput, 255)
                .Parameters.Append .CreateParameter("mRptFile", adVarBinary, adParamInput, FileStream.Size)
                .Parameters.Append .CreateParameter("mRptUser", adVarChar, adParamInput, 255)
                .Parameters.Append .CreateParameter("mRptFileSize", adBigInt, adParamInput)
                .Parameters("mRptName") = rName
                .Parameters("mRptDesc") = rDesc
                .Parameters("mRptSectionCode") = rScode
                .Parameters("mRptFile") = myFile 'FileStream.Read(FileStream.Size)
                .Parameters("mRptUser") = mUserInit
                If myFile(UBound(myFile)) = 0 Then
                    .Parameters("mRptFileSize") = myFile(UBound(myFile))
                Else
                    .Parameters("mRptFileSize") = Null
                End If
                .Execute
                If PromptMessage Then MsgBox "Report has been successfully saved.", vbInformation + vbOKOnly, "Reports"
                Saved = True
            End With
        Else
            MsgBox "Please input your initial"
        End If
    End If
Else
    If Len(Trim(rUser)) = 0 Then
        mUserInit = mClass.showInputBox("Input your initial then press Enter.", , , , 10, True, False, False)
    Else
        mUserInit = rUser
    End If
    If Len(mUserInit) > 0 And lblSize.Caption <> "0K" Then
        myFile = FileStream.Read(FileStream.Size)
        With mycommand
            .ActiveConnection = myConnection
            .CommandType = adCmdStoredProc
            .CommandText = "sp_ReportUpdate"
            .Parameters.Append .CreateParameter("mRptID", adInteger, adParamInput)
            .Parameters.Append .CreateParameter("mRptName", adVarChar, adParamInput, 255)
            .Parameters.Append .CreateParameter("mRptDesc", adVarChar, adParamInput, 255)
            .Parameters.Append .CreateParameter("mRptFile", adVarWChar, adParamInput, FileStream.Size)
            .Parameters.Append .CreateParameter("mRptUser", adVarChar, adParamInput, 255)
            .Parameters.Append .CreateParameter("mRptFileSize", adBigInt, adParamInput)
            .Parameters("mRptID") = rID
            .Parameters("mRptName") = rName
            .Parameters("mRptDesc") = rDesc
            .Parameters("mRptFile") = myFile 'FileStream.Read(FileStream.Size)
            .Parameters("mRptUser") = mUserInit
            If myFile(UBound(myFile)) = 0 Then
                .Parameters("mRptFileSize") = myFile(UBound(myFile))
            Else
                .Parameters("mRptFileSize") = Null
            End If
            .Execute
            If PromptMessage Then MsgBox "Report has been successfully Updated.", vbInformation + vbOKOnly, "Reports"
            Saved = True
        End With
    Else
        MsgBox "Please input your initial"
    End If

End If
SaveReport = Saved
End Function

Private Sub ClearValues()
    If Len(Trim(txtReportName.Text)) > 0 Then
        If Len(Trim(Dir(IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\" & txtReportName.Text))) > 0 Then
            Kill IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\" & txtReportName.Text
        End If
    End If
    lblReportID.Caption = ""
    txtReportName.Text = ""
    txtExtName.Text = ""
    lblSize.Caption = "0K"
    txtDisplayName.Text = ""
    txtNote.Text = ""
    lblModDate.Caption = ""
    lblModBy.Caption = ""
    cmdCancel.Enabled = False
    cmdUpdate.Enabled = False
    msfgResult.Enabled = True
    cmdSave.Enabled = True
    cmdUpload.Enabled = False
    cmdClose.Cancel = True
End Sub

Private Function HasReport(mReportName As String) As Boolean
Dim rstTemp As ADODB.Recordset
Dim mhasReport As Boolean
mhasReport = False
Set rstTemp = myConnection.Execute("Select * from tbLabReportNames where rptName = '" & mReportName & "'")
If rstTemp.RecordCount > 0 Then
    mhasReport = True
End If
Set rstTemp = Nothing
HasReport = mhasReport
End Function

Private Sub SetGridDefault(mRow As Integer, Optional Reload As Boolean = False)
Dim ctr As Integer
Dim cRow As Integer
With msfgResult
    cRow = .Row
    If Reload Then
        For ctr = 1 To .rows - 1
            If .TextMatrix(ctr, 4) = "True" Then
                .Row = ctr
                .Col = 2
                .CellFontBold = True
            End If
        Next ctr
    Else
        .Row = mRow
        .Col = 2
        .CellFontBold = True
    End If
    .Row = cRow
End With
End Sub

Private Sub assignValues()
Dim rstTemp As ADODB.Recordset
Dim tempSize As Long
Dim mFile() As Byte
With msfgResult
    lblReportID.Caption = .TextMatrix(.Row, 1)
    txtDisplayName.Text = .TextMatrix(.Row, 2)
    txtExtName.Text = "RPT"
    txtNote.Text = .TextMatrix(.Row, 3)
    lblModDate.Caption = .TextMatrix(.Row, 6)
    lblModBy.Caption = .TextMatrix(.Row, 7)
    Set rstTemp = myConnection.Execute("Select rptFile, rptFileSize from tbLabReportNames where rptID = " & .TextMatrix(.Row, 1))
    If rstTemp.RecordCount > 0 Then
        rstTemp.MoveFirst
        If FileStream.State = adStateOpen Then FileStream.Close
        With FileStream
            .Type = adTypeBinary
            .Open
            .Flush
            mFile = rstTemp.Fields("RptFile").Value & ""
            If Len(rstTemp.Fields("RptFile").Value & "") > 0 Then
                If CStr(rstTemp.Fields("RptFileSize").Value & "") = "0" Then
                    If mFile(UBound(mFile)) <> rstTemp.Fields("RptFileSize").Value Then
                        ReDim Preserve mFile(UBound(mFile) + 1)
                        mFile(UBound(mFile)) = 0
                    End If
                End If
            End If
            .Write mFile
            '.Write rstTemp.Fields("RptFile").Value
'            .SaveToFile IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\" & txtDisplayName.Text, adSaveCreateOverWrite
'            .LoadFromFile IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\" & txtDisplayName.Text
'            Kill IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\" & txtDisplayName.Text
            lblSize.Caption = CStr(CInt((.Size / 1024))) & "KB"
            If .Size > 0 Then
                cmdEdit.Enabled = True
                cmdUpload.Enabled = False
            End If
        End With
    End If
End With
Set rstTemp = Nothing
End Sub
