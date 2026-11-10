VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmStatistics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Statistics"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   9495
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   10830
      TabIndex        =   0
      Top             =   2700
      Width           =   1665
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Statistics [Beta]"
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
         Left            =   90
         TabIndex        =   1
         Top             =   150
         Width           =   1605
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   6315
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   9435
      Begin VB.Frame fraBody 
         Appearance      =   0  'Flat
         BackColor       =   &H006B4B01&
         ForeColor       =   &H80000008&
         Height          =   1515
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   4110
         Width           =   9135
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5850
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   1140
            Width           =   2565
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5850
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   810
            Width           =   2565
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5850
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   150
            Width           =   2565
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5850
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   480
            Width           =   2565
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1410
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   1140
            Width           =   1575
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1410
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   810
            Width           =   1575
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1410
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1410
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   150
            Width           =   1575
         End
         Begin VB.Label lblText 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Average samples (per running day):"
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
            Index           =   5
            Left            =   3150
            TabIndex        =   21
            Top             =   1170
            Width           =   2625
         End
         Begin VB.Label lblText 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Average samples (per day):"
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
            Left            =   3720
            TabIndex        =   19
            Top             =   810
            Width           =   2040
         End
         Begin VB.Label lblText 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date with least samples:"
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
            Left            =   3960
            TabIndex        =   17
            Top             =   510
            Width           =   1770
         End
         Begin VB.Label lblText 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date with most samples:"
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
            Index           =   6
            Left            =   3960
            TabIndex        =   16
            Top             =   210
            Width           =   1770
         End
         Begin VB.Label lblText 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unedited:"
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
            Index           =   3
            Left            =   660
            TabIndex        =   9
            Top             =   870
            Width           =   675
         End
         Begin VB.Label lblText 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Samples:"
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
            Index           =   0
            Left            =   330
            TabIndex        =   8
            Top             =   1200
            Width           =   1050
         End
         Begin VB.Label lblText 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Printed Results:"
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
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   210
            Width           =   1125
         End
         Begin VB.Label lblText 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saved:"
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
            Index           =   2
            Left            =   840
            TabIndex        =   6
            Top             =   540
            Width           =   510
         End
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
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
         Height          =   465
         Left            =   7920
         TabIndex        =   3
         Top             =   5730
         Width           =   1305
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   3825
         Left            =   120
         TabIndex        =   4
         Top             =   270
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   6747
         _Version        =   393216
         BackColor       =   16777215
         RowHeightMin    =   360
         BackColorFixed  =   10906411
         ForeColorFixed  =   16777215
         BackColorBkg    =   12632256
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
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
End
Attribute VB_Name = "frmStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim rstTemp As ADODB.Recordset
    setFormColors Me
    setmsfgResult
    ReturnValue = ""
    'All Samples
    'Set rstTemp = myConnection.Execute("Select format(resultdate,""mm/dd/yyyy"") as [Date], count(format(resultdate,""mm/dd/yyyy"")) as [Count]  from (select distinct resultdate from tblabinstrumentresults) group by format(resultdate,""mm/dd/yyyy"")")
    'Printed Samples
    Set rstTemp = myConnection.Execute("Select format(resultdate,""mm/dd/yyyy"") as [Date], count(format(resultdate,""dddd mm/dd/yyyy"")) as [Count]  from (select distinct resultdate from tblabinstrumentresults where instrumentid = '" & InstrumentID & "') group by format(resultdate,""mm/dd/yyyy"")")
    Set msfgResult.DataSource = rstTemp
    countData
End Sub

Private Sub setmsfgResult()
With msfgResult
    If .Rows >= 15 Then
        .ColWidth(0) = 0
        .ColWidth(1) = (.Width - 300) * 0.3
        .ColWidth(2) = (.Width - 300) * 0.3
    Else
        .ColWidth(0) = 400
        .ColWidth(1) = (.Width - 400) * 0.4
        .ColWidth(2) = (.Width - 400) * 0.4
    End If
    .ColWidth(3) = 0
End With
End Sub

Private Sub countData()
Dim rstTemp As ADODB.Recordset
    Set rstTemp = myConnection.Execute("Select distinct resultdate, instrumentid, accessionnum  from tblabinstrumentresults where printed = TRUE")
    Text1.Text = rstTemp.RecordCount
    Set rstTemp = myConnection.Execute("Select distinct resultdate, instrumentid, accessionnum  from tblabinstrumentresults where printed = FALSE and len(refnum) > 0")
    Text2.Text = rstTemp.RecordCount
    Set rstTemp = myConnection.Execute("Select distinct resultdate, instrumentid, accessionnum  from tblabinstrumentresults where printed = FALSE and (len(refnum) = 0 or isnull(refnum))")
    Text3.Text = rstTemp.RecordCount
    Set rstTemp = myConnection.Execute("Select distinct resultdate, instrumentid, accessionnum  from tblabinstrumentresults")
    Text4.Text = rstTemp.RecordCount
    Set rstTemp = myConnection.Execute("Select format(resultdate,""mm/dd/yyyy"") as [Date], count(format(resultdate,""dddd mm/dd/yyyy"")) as [Count]  from (select distinct resultdate from tblabinstrumentresults) group by format(resultdate,""mm/dd/yyyy"") order by count(format(resultdate,""dddd mm/dd/yyyy""))")
    If rstTemp.RecordCount > 0 Then
        rstTemp.MoveLast
        Text8.Text = rstTemp.Fields("Date").Value & " @ " & rstTemp.Fields("Count").Value
        rstTemp.MoveFirst
        Text7.Text = rstTemp.Fields("Date").Value & " @ " & rstTemp.Fields("Count").Value
        Text5.Text = Format(getAverage, "0")
        Set rstTemp = myConnection.Execute("Select min(resultdate) as oldest from tblabinstrumentresults")
        Text6.Text = Format(Val(Text4.Text) / Val(DateDiff("d", rstTemp.Fields("oldest").Value, Now)), "0")
    End If
End Sub

Private Function getAverage() As String
Dim ctr As Integer
Dim mTotal As Long
mTotal = 0
With msfgResult
    For ctr = 1 To .Rows - 1
        mTotal = mTotal + .TextMatrix(ctr, 2)
    Next ctr
    If mTotal > 0 Then getAverage = mTotal / (.Rows - 1) Else getAverage = "0"
End With
End Function

Private Sub msfgResult_DblClick()
ReturnValue = msfgResult.TextMatrix(msfgResult.Row, 1)
Unload Me
End Sub
