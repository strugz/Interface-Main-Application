VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLabExam 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Exams"
   ClientHeight    =   8670
   ClientLeft      =   1815
   ClientTop       =   2085
   ClientWidth     =   11970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdUp 
      Height          =   975
      Left            =   12000
      Picture         =   "frmLabExam.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2850
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton cmdDown 
      Height          =   975
      Left            =   12000
      Picture         =   "frmLabExam.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   3870
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   12690
      TabIndex        =   13
      Top             =   1440
      Width           =   3015
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Build File - [Laboratory Exam]"
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
         TabIndex        =   14
         Top             =   150
         Width           =   2850
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   8595
      Index           =   1
      Left            =   3810
      TabIndex        =   24
      Top             =   30
      Width           =   8115
      Begin VB.CommandButton Command2 
         Caption         =   "Save Sort Order"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5130
         TabIndex        =   48
         Top             =   8130
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.CommandButton cmdGenerateSort 
         Caption         =   "Generate Sort Order"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   47
         Top             =   8130
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6630
         TabIndex        =   26
         Top             =   8130
         Width           =   1395
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   7785
         Left            =   60
         TabIndex        =   25
         Top             =   240
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   13732
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
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Double Click an Item to Edit."
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
         Left            =   240
         TabIndex        =   27
         Top             =   8280
         Width           =   2280
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H00C0C0C0&
      Height          =   8595
      Index           =   0
      Left            =   30
      TabIndex        =   15
      Top             =   30
      Width           =   3735
      Begin VB.TextBox txtDefaultFormat 
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
         Left            =   1260
         TabIndex        =   46
         Top             =   5850
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.TextBox txtDefaultFactor 
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
         Left            =   1260
         TabIndex        =   45
         Top             =   5520
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.ComboBox cboMVfield 
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
         Height          =   330
         ItemData        =   "frmLabExam.frx":1404
         Left            =   1080
         List            =   "frmLabExam.frx":1411
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   5160
         Width           =   2445
      End
      Begin VB.ComboBox cboResultType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmLabExam.frx":1421
         Left            =   1080
         List            =   "frmLabExam.frx":142E
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   4770
         Width           =   2445
      End
      Begin RichTextLib.RichTextBox txtExamNotes 
         Height          =   2115
         Left            =   90
         TabIndex        =   36
         Top             =   5940
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   3731
         _Version        =   393217
         MaxLength       =   255
         Appearance      =   0
         TextRTF         =   $"frmLabExam.frx":143E
      End
      Begin VB.TextBox txtHiloFlags 
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
         TabIndex        =   33
         Top             =   4410
         Width           =   2565
      End
      Begin VB.TextBox txtFormula 
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
         Left            =   780
         TabIndex        =   31
         Top             =   4080
         Width           =   2775
      End
      Begin VB.TextBox txtDilution 
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
         Left            =   750
         MaxLength       =   8
         TabIndex        =   8
         Top             =   3420
         Width           =   2805
      End
      Begin VB.TextBox txtFactorOP 
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
         Left            =   1410
         MaxLength       =   8
         TabIndex        =   9
         Top             =   3750
         Width           =   2145
      End
      Begin VB.ComboBox cboDiff 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmLabExam.frx":14C0
         Left            =   1020
         List            =   "frmLabExam.frx":14CA
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3060
         Width           =   2535
      End
      Begin VB.TextBox txtUnit 
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
         Left            =   570
         MaxLength       =   8
         TabIndex        =   2
         Top             =   1320
         Width           =   3015
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   630
         TabIndex        =   11
         Top             =   8100
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CommandButton cmdAddSave 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   8100
         Width           =   1485
      End
      Begin VB.TextBox txtResultCode 
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
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   5
         Top             =   2310
         Width           =   2265
      End
      Begin VB.TextBox txtInstrumentExamName 
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
         Left            =   1830
         MaxLength       =   15
         TabIndex        =   1
         Top             =   990
         Width           =   1755
      End
      Begin VB.TextBox txtLISExamCode 
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
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1980
         Width           =   2265
      End
      Begin VB.TextBox txtDisplayName 
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
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   3
         Top             =   1650
         Width           =   2265
      End
      Begin VB.TextBox txtRefNum 
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
         Left            =   1830
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   660
         Width           =   1755
      End
      Begin VB.TextBox txtInstrumentID 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   1170
         MaxLength       =   15
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "HMX"
         Top             =   330
         Width           =   2415
      End
      Begin VB.ComboBox cboActive 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmLabExam.frx":14DB
         Left            =   750
         List            =   "frmLabExam.frx":14E5
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2670
         Width           =   2805
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Default Format:"
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
         Index           =   17
         Left            =   120
         TabIndex        =   44
         Top             =   5910
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Default Factor:"
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
         Index           =   16
         Left            =   120
         TabIndex        =   43
         Top             =   5550
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "MV Field:"
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
         Index           =   15
         Left            =   330
         TabIndex        =   39
         Top             =   5190
         Width           =   660
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Result Type:"
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
         Index           =   14
         Left            =   120
         TabIndex        =   37
         Top             =   4800
         Width           =   900
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Notes:"
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
         Index           =   13
         Left            =   120
         TabIndex        =   35
         Top             =   5730
         Width           =   900
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Hi-Lo Flags:"
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
         Index           =   12
         Left            =   120
         TabIndex        =   34
         Top             =   4470
         Width           =   855
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Formula:"
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
         Index           =   11
         Left            =   120
         TabIndex        =   32
         Top             =   4140
         Width           =   615
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Factor:"
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
         Index           =   9
         Left            =   120
         TabIndex        =   30
         Top             =   3450
         Width           =   510
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Factor Operator:"
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
         Index           =   10
         Left            =   120
         TabIndex        =   29
         Top             =   3810
         Width           =   1200
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Diff Count:"
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
         Left            =   120
         TabIndex        =   28
         Top             =   3120
         Width           =   765
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit:"
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
         Left            =   180
         TabIndex        =   23
         Top             =   1350
         Width           =   315
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "LIS Result Code:"
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
         Left            =   150
         TabIndex        =   22
         Top             =   2400
         Width           =   1185
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Active:"
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
         Left            =   120
         TabIndex        =   21
         Top             =   2730
         Width           =   510
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "LIS Exam Code:"
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
         Left            =   150
         TabIndex        =   20
         Top             =   2010
         Width           =   1125
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Display Name:"
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
         Left            =   150
         TabIndex        =   19
         Top             =   1710
         Width           =   1020
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument Exam Name:"
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
         Left            =   150
         TabIndex        =   18
         Top             =   1020
         Width           =   1680
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument Exam Code:"
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
         Left            =   180
         TabIndex        =   17
         Top             =   690
         Width           =   1650
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Instrument ID:"
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
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmLabExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstExams As adodb.Recordset
Private AddMode As Boolean
Private mHISexamcode As String
Private mHISresultcode As String
Private mInstrumentID As String
Private mForms As New ClassForms

Private Sub cboActive_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub cboResultType_Click()
        If cboResultType.ListIndex = 2 Then
            cboMVfield.Enabled = True
            fillCboMVfield
        Else
            cboMVfield.Enabled = False
            cboMVfield.Clear
        End If
End Sub

Private Sub cmdAddSave_Click()
Dim rstCheck As adodb.Recordset
Dim rstSaveData As adodb.Recordset
If AddMode Then
    If Len(txtResultCode.Text) = 0 Or Len(txtInstrumentExamName.Text) = 0 Or Len(txtLISExamCode.Text) = 0 Then
        MsgBox "Unable to save data.", vbOKOnly + vbInformation, "Data Check"
        txtInstrumentExamName.SetFocus
    Else
        Set rstCheck = myConnection.Execute("qryCheckExam '" & txtInstrumentID.Text & "','" & txtInstrumentExamName.Text & "','" & txtLISExamCode.Text & "','" & txtResultCode.Text & "'")
        If Len(Trim(txtDilution.Text)) = 0 Then txtDilution.Text = 1
        If rstCheck.RecordCount = 0 Then
            Set rstSaveData = myConnection.Execute("qryInsertExam '" & txtInstrumentID.Text & "','" & txtRefNum.Text & "','" & txtInstrumentExamName.Text & "','" & txtLISExamCode.Text & "','" & Replace(txtDisplayName.Text, "'", "''") & "'," & cboActive.Text & ",'" & txtResultCode.Text & "','" & txtUnit.Text & "'," & cboDiff.Text & ",'" & txtDilution.Text & "','" & txtFactorOP.Text & "','" & txtFormula.Text & "','" & txtHiloFlags.Text & "','" & txtExamNotes.Text & "','" & cboResultType.Text & "','" & cboMVfield.Text & "'," & Val(txtDefaultFactor.Text) & ",'" & txtDefaultFormat.Text & "'")
            ClearValues
            txtInstrumentExamName.SetFocus
            msfgResult.Refresh
        Else
            If txtLISExamCode.Text = "0" And txtResultCode.Text = "0" Then
                Set rstSaveData = myConnection.Execute("qryInsertExam '" & txtInstrumentID.Text & "','" & txtRefNum.Text & "','" & txtInstrumentExamName.Text & "','" & txtLISExamCode.Text & "','" & Replace(txtDisplayName.Text, "'", "''") & "'," & cboActive.Text & ",'" & txtResultCode.Text & "','" & txtUnit.Text & "'," & cboDiff.Text & ",'" & txtDilution.Text & "','" & txtFactorOP.Text & "','" & txtFormula.Text & "','" & txtHiloFlags.Text & "','" & txtExamNotes.Text & "','" & cboResultType.Text & "','" & cboMVfield.Text & "'," & txtDefaultFactor.Text & ",'" & txtDefaultFormat.Text & "'")
                ClearValues
                txtInstrumentExamName.SetFocus
                msfgResult.Refresh
            Else
                MsgBox "Record Already Exists.", vbOKOnly + vbInformation, "Record Check"
            End If
        End If
    End If
Else
    Set rstCheck = myConnection.Execute("qryCheckExam2 '" & txtInstrumentID.Text & "','" & txtRefNum.Text & "','" & txtLISExamCode.Text & "','" & txtResultCode.Text & "'")
    If Len(Trim(txtDilution.Text)) = 0 Then txtDilution.Text = 1
    If rstCheck.RecordCount = 0 Then
        Set rstSaveData = myConnection.Execute("qryUpdateExam '" & txtInstrumentID.Text & "','" & txtRefNum.Text & "','" & txtInstrumentExamName.Text & "','" & txtLISExamCode.Text & "','" & Replace(txtDisplayName.Text, "'", "''") & "'," & IIf(cboActive.Text = "TRUE", "1", "0") & ",'" & txtResultCode.Text & "','" & txtUnit.Text & "'," & IIf(cboDiff.Text = "TRUE", "1", "0") & ",'" & txtDilution.Text & "','" & txtFactorOP.Text & "','" & txtFormula.Text & "','" & txtHiloFlags.Text & "','" & txtExamNotes.Text & "','" & cboResultType.Text & "','" & cboMVfield.Text & "','" & txtDefaultFactor.Text & "','" & txtDefaultFormat.Text & "'")
        cmdCancel.Visible = False
        cmdAddSave.Caption = "Add"
        AddMode = True
        msfgResult.Enabled = True
        ClearValues
        cmdClose.Cancel = True
        lblStatus.Caption = "Double Click an Item to Edit."
    Else
        If txtLISExamCode.Text = "0" And txtResultCode.Text = "0" Then
            Set rstSaveData = myConnection.Execute("qryUpdateExam '" & txtInstrumentID.Text & "','" & txtRefNum.Text & "','" & txtInstrumentExamName.Text & "','" & txtLISExamCode.Text & "','" & Replace(txtDisplayName.Text, "'", "''") & "'," & cboActive.Text & ",'" & txtResultCode.Text & "','" & txtUnit.Text & "'," & cboDiff.Text & ",'" & txtDilution.Text & "','" & txtFactorOP.Text & "','" & txtFormula.Text & "','" & txtHiloFlags.Text & "','" & txtExamNotes.Text & "','" & cboResultType.Text & "','" & cboMVfield.Text & "'," & IIf(Len(txtDefaultFactor.Text) > 0, Val(txtDefaultFactor.Text), "NULL") & ",'" & txtDefaultFormat.Text & "'")
            cmdCancel.Visible = False
            cmdAddSave.Caption = "Add"
            AddMode = True
            msfgResult.Enabled = True
            ClearValues
            cmdClose.Cancel = True
            lblStatus.Caption = "Double Click an Item to Edit."
        Else
            MsgBox "Record Already Exists.", vbOKOnly + vbInformation, "Record Check"
        End If
    End If
End If
setValues
End Sub

Private Sub cmdCancel_Click()
cmdCancel.Visible = False
cmdAddSave.Caption = "Add"
AddMode = True
msfgResult.Enabled = True
ClearValues
cmdClose.Cancel = True
lblStatus.Caption = "Double Click an Item to Edit."
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDown_Click()
Dim ctr As Integer
Dim tempStr(14) As String
With msfgResult
    For ctr = 1 To .Cols - 1
        tempStr(ctr) = .TextMatrix(.Row, ctr)
    Next ctr
    For ctr = 1 To .Cols - 1
        .TextMatrix(.Row, ctr) = .TextMatrix(.Row + 1, ctr)
    Next ctr
    For ctr = 1 To .Cols - 1
        .TextMatrix(.Row + 1, ctr) = tempStr(ctr)
    Next ctr
    .Row = .Row + 1
    .ColSel = 13
    msfgResult_EnterCell
End With
cmdGenerateSort_Click
End Sub

Private Sub cmdGenerateSort_Click()
Dim ctr As Integer
With msfgResult
    For ctr = 1 To .rows - 1
        .TextMatrix(ctr, 13) = ctr
    Next ctr
End With
End Sub

Private Sub cmdUp_Click()
Dim ctr As Integer
Dim tempStr(14) As String
With msfgResult
    For ctr = 0 To .Cols - 1
        tempStr(ctr) = .TextMatrix(.Row, ctr)
    Next ctr
    For ctr = 0 To .Cols - 1
        .TextMatrix(.Row, ctr) = .TextMatrix(.Row - 1, ctr)
    Next ctr
    For ctr = 0 To .Cols - 1
        .TextMatrix(.Row - 1, ctr) = tempStr(ctr)
    Next ctr
    .Row = .Row - 1
    .ColSel = 13
    msfgResult_EnterCell
End With
cmdGenerateSort_Click
End Sub

Private Sub Command2_Click()
'Dim ctr As Integer
'With msfgResult
'    For ctr = 1 To .Rows - 1
'        myConnection.Execute "update tbcolabexam set defaultsortnum = " & .TextMatrix(ctr, 13) & " where instrumentid = '" & mInstrumentID & "' and examid = '" & .TextMatrix(ctr, 0) & "'"
'    Next ctr
'End With
'MsgBox "Sort Order Saved.", vbOKOnly + vbInformation
End Sub

Private Sub Form_Load()
setFormColors Me
    mInstrumentID = InstrumentID 'mForms.SelectInstrument
    setValues
    AddMode = True
End Sub

Private Function getRefNum() As String
Dim rstData As adodb.Recordset
    If DBServerType = "LOCAL DATABASE" Then
        Set rstData = myConnection.Execute("Select max(cint(ExamID)) as refnum from tbCoLabExam")
    ElseIf DBServerType = "MYSQL SERVER" Then
        Set rstData = myConnection.Execute("Select max(cast(ExamID as signed)) ""refnum"" from tbCoLabExam")
    ElseIf DBServerType = "MSSQL SERVER" Then
        Set rstData = myConnection.Execute("Select max(cast(ExamID as int)) ""refnum"" from tbCoLabExam")
    End If
    If rstData.State = 1 Then
        If rstData.RecordCount > 0 Then
            getRefNum = IIf(IsNull(rstData!refnum), "0", rstData!refnum)
            getRefNum = Val(getRefNum) + 1
        Else
            getRefNum = "1"
        End If
    End If
End Function

Private Sub setValues()
Dim rstHIScodes As adodb.Recordset
Dim rstHIScodes2 As adodb.Recordset
    If DBServerType = "LOCAL DATABASE" Then
        Set rstHIScodes = myConnection.Execute("Select max(int('0' + HisExamCode)) as ExamCode, max(int('0' + HisResultCode)) as ResultCode from tbCoLabExam where instrumentid = '" & mInstrumentID & "'")
    ElseIf DBServerType = "MYSQL SERVER" Then
        Set rstHIScodes = myConnection.Execute("Select max(cast(HisExamCode as signed)) ""ExamCode"", max(cast(HisResultCode as signed)) ""ResultCode"" from tbCoLabExam where instrumentid = '" & mInstrumentID & "' order by defaultsortnum")
    ElseIf DBServerType = "MSSQL SERVER" Then
        Set rstHIScodes = myConnection.Execute("Select max(cast(HisExamCode as int)) ""ExamCode"", max(cast(HisResultCode as int)) ""ResultCode"" from tbCoLabExam where instrumentid = '" & mInstrumentID & "'")
    End If
    If Len(Trim(rstHIScodes.Fields("ExamCode").Value) & "") > 0 Then
        mHISexamcode = Trim(rstHIScodes.Fields("ExamCode").Value) & ""
    Else
        Set rstHIScodes2 = myConnection.Execute("Select max(HisExamCode) as ExamCode from tbCoLabExam")
        mHISexamcode = Trim(str(Val(Trim(rstHIScodes2.Fields("ExamCode").Value) & "") + 1))
    End If
    
    If Len(Trim(rstHIScodes.Fields("ResultCode").Value) & "") > 0 Then
        mHISresultcode = Trim(str(Val(Trim(rstHIScodes.Fields("ResultCode").Value)) + 1))
    Else
        mHISresultcode = "1"
    End If

    txtInstrumentID.Text = mInstrumentID
    txtRefNum.Text = getRefNum
    txtLISExamCode.Text = mHISexamcode
    txtResultCode.Text = mHISresultcode
    cboActive.ListIndex = 0
    cboDiff.ListIndex = 1
    cboResultType.ListIndex = 0
    If DBServerType = "LOCAL DATABASE" Then
        Set rstExams = myConnection.Execute("qrySelectExams '" & txtInstrumentID.Text & "'")
    ElseIf DBServerType = "MYSQL SERVER" Then
        Set rstExams = myConnection.Execute("call qrySelectExams ('" & txtInstrumentID.Text & "');")
    ElseIf DBServerType = "MSSQL SERVER" Then
        Set rstExams = myConnection.Execute("exec qrySelectExams '" & txtInstrumentID.Text & "';")
    End If
    If rstExams.RecordCount > 0 Then
        SetGrid False
        Set msfgResult.Recordset = rstExams
    Else
        SetGrid
    End If
End Sub

Private Sub setItems(ItemId As String)
Dim rstItems As New adodb.Recordset
If Len(Trim(ItemId)) > 0 Then
    Set rstItems = myConnection.Execute("qryGetExamItem " & ItemId)
    With rstItems
        txtRefNum.Text = !ExamID
        txtInstrumentExamName.Text = !InstExamName & ""
        txtDisplayName.Text = !DisplayName & ""
        If Len(Trim(!HisExamCode & "")) > 0 Then txtLISExamCode.Text = !HisExamCode & "" Else txtLISExamCode.Text = mHISexamcode
        If Len(Trim(!HISresultcode & "")) > 0 Then txtResultCode.Text = !HISresultcode & "" Else txtResultCode.Text = mHISresultcode
        txtUnit.Text = !Unit & ""
        If !Active Then cboActive.ListIndex = 0 Else cboActive.ListIndex = 1
        If !DiffCount Then cboDiff.ListIndex = 0 Else cboDiff.ListIndex = 1
        txtDilution.Text = rstItems.Fields("Dilution").Value & ""
        txtFactorOP.Text = rstItems.Fields("FactorOP").Value & ""
        txtFormula.Text = rstItems.Fields("Formula").Value & ""
        txtHiloFlags.Text = rstItems.Fields("HiloFlags").Value & ""
        If txtLISExamCode.Text <> "0" And txtResultCode.Text <> "0" Then txtFormula.Enabled = False Else txtFormula.Enabled = True
        txtExamNotes.Text = rstItems.Fields("ExamNotes").Value & ""
        If !ExamType = "NV" Then
            cboResultType.ListIndex = 0
        ElseIf !ExamType = "SV" Then
            cboResultType.ListIndex = 1
        ElseIf !ExamType = "MV" Then
            cboResultType.ListIndex = 2
        End If
        If cboResultType.ListIndex = 2 Then
            cboMVfield.Enabled = True
            fillCboMVfield
            cboMVfield.ListIndex = findCboIndex(cboMVfield, !ExamTypeValues & "")
        Else
            cboMVfield.Enabled = False
            cboMVfield.Clear
        End If
'        txtDefaultFactor.Text = rstItems.Fields("DefaultFactor").Value & ""
'        txtDefaultFormat.Text = rstItems.Fields("DefaultDataFormat").Value & ""
    End With
End If
End Sub

Private Sub fillCboMVfield()
Dim rstTypes As adodb.Recordset
Set rstTypes = myConnection.Execute("Select * from tbSelectionType where status = '1'")
If rstTypes.RecordCount > 0 Then
    rstTypes.MoveFirst
    cboMVfield.Clear
    Do While Not rstTypes.EOF
        cboMVfield.AddItem rstTypes.Fields("SelType").Value
        rstTypes.MoveNext
    Loop
End If
End Sub


Private Sub ClearValues()
    txtRefNum.Text = getRefNum
    txtInstrumentExamName.Text = ""
    txtDisplayName.Text = ""
    txtLISExamCode.Text = mHISexamcode
    txtResultCode.Text = mHISresultcode
    txtUnit.Text = ""
    txtDilution.Text = ""
    txtFactorOP.Text = ""
    cboActive.ListIndex = 0
    cboDiff.ListIndex = 1
    txtFormula.Text = ""
    txtHiloFlags.Text = ""
    txtExamNotes.Text = ""
    cboResultType.ListIndex = 0
    txtDefaultFactor.Text = ""
    txtDefaultFormat.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rstExams = Nothing
End Sub

Private Sub msfgResult_Click()
Dim numSort As Integer

If msfgResult.MouseRow = 0 Then
    msfgResult.Col = msfgResult.MouseCol
    
    If numSort Then
        msfgResult.Sort = flexSortGenericDescending
    Else
        msfgResult.Sort = flexSortGenericAscending
    End If
End If
End Sub

Private Sub msfgResult_DblClick()
With msfgResult
    If Len(Trim(.TextMatrix(.Row, 0))) > 0 Then
        setItems .TextMatrix(.Row, 0)
        AddMode = False
        cmdAddSave.Caption = "Update"
        cmdCancel.Visible = True
        msfgResult.Enabled = False
        cmdCancel.Cancel = True
        txtInstrumentExamName.SetFocus
        lblStatus.Caption = "Editing Instrument Exam Number: " & .TextMatrix(.Row, 0)
    End If
End With
End Sub

Private Sub msfgResult_EnterCell()
If msfgResult.Row = 1 Then
    cmdUp.Enabled = False
    cmdDown.Enabled = True
ElseIf msfgResult.Row = msfgResult.rows - 1 Then
    cmdUp.Enabled = True
    cmdDown.Enabled = False
Else
    cmdUp.Enabled = True
    cmdDown.Enabled = True
End If
End Sub

Private Sub txtDisplayName_GotFocus()
highLight
End Sub

Private Sub txtDisplayName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtInstrumentExamName_GotFocus()
highLight
End Sub

Private Sub txtInstrumentExamName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtInstrumentExamName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtLISExamCode_GotFocus()
highLight
End Sub

Private Sub txtLISExamCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub SetGrid(Optional isclear As Boolean = True)
With msfgResult
If isclear Then
    .Clear
    .Cols = 7
    .rows = 2
    .TextMatrix(0, 0) = "Code"
    .TextMatrix(0, 1) = "Exam"
    .TextMatrix(0, 2) = "Instrument Exam"
    .TextMatrix(0, 3) = "Exam Code"
    .TextMatrix(0, 4) = "Result Code"
    .TextMatrix(0, 5) = "Active"
    .TextMatrix(0, 6) = "Diff"
Else

End If
    .ColWidth(0) = 500
    .ColWidth(1) = 2400
    .ColWidth(2) = 1300
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 800
    .ColWidth(6) = 600
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
    .ColWidth(10) = 1000
    .ColWidth(11) = 1000
End With
End Sub

Private Sub txtResultCode_GotFocus()
    highLight
End Sub

Private Sub txtResultCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtUnit_GotFocus()
    highLight
End Sub

Private Sub txtUnit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Function findCboIndex(cboTemp As ComboBox, texttofind As String) As Integer
Dim ctr As Integer
With cboTemp
    For ctr = 0 To .ListCount - 1
        .ListIndex = ctr
        If .Text = texttofind Then
            findCboIndex = .ListIndex
            Exit Function
        End If
    Next ctr
    .ListIndex = -1
End With
End Function

