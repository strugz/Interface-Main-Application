VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWorkList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Requests"
   ClientHeight    =   10320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13395
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   13395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRemoveFromWorklist 
      Caption         =   "Remove From Instrument Worklist"
      Height          =   435
      Left            =   7320
      TabIndex        =   72
      Top             =   9750
      Width           =   2745
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   10000
      Left            =   9390
      Top             =   9780
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   465
      Left            =   10140
      TabIndex        =   8
      Top             =   9720
      Width           =   1455
   End
   Begin VB.CommandButton cmdPos 
      Caption         =   "Set Cassette Position"
      Height          =   345
      Left            =   3240
      TabIndex        =   6
      Top             =   9810
      Width           =   1965
   End
   Begin VB.TextBox txtCass 
      Height          =   285
      Left            =   2220
      TabIndex        =   5
      Top             =   9840
      Width           =   975
   End
   Begin VB.TextBox txtRack 
      Height          =   285
      Left            =   570
      TabIndex        =   4
      Top             =   9810
      Width           =   975
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send/Resend Request"
      Height          =   435
      Left            =   5280
      TabIndex        =   7
      Top             =   9750
      Width           =   2025
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   465
      Left            =   11670
      TabIndex        =   9
      Top             =   9720
      Width           =   1545
   End
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   14010
      TabIndex        =   10
      Top             =   180
      Width           =   11295
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Worklist"
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
         TabIndex        =   11
         Top             =   150
         Width           =   795
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   9585
      Index           =   0
      Left            =   30
      TabIndex        =   12
      Top             =   60
      Width           =   13215
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Left            =   11040
         TabIndex        =   73
         Top             =   300
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Format          =   55836673
         CurrentDate     =   40143
      End
      Begin VB.TextBox txtExams 
         Height          =   345
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   9150
         Width           =   12855
      End
      Begin VB.ComboBox cboFilter 
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
         Left            =   7260
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   2535
      End
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   0
         Top             =   270
         Width           =   5745
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   8295
         Left            =   150
         TabIndex        =   2
         Top             =   780
         Width           =   12885
         _ExtentX        =   22728
         _ExtentY        =   14631
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
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Request Date:"
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
         Height          =   195
         Left            =   9960
         TabIndex        =   74
         Top             =   330
         Width           =   1185
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Filter:"
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
         Height          =   195
         Left            =   6780
         TabIndex        =   69
         Top             =   330
         Width           =   675
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Search:"
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
         Height          =   195
         Left            =   240
         TabIndex        =   68
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   2145
      Index           =   1
      Left            =   13470
      TabIndex        =   13
      Top             =   9360
      Width           =   11295
      Begin VB.Frame fraBody 
         BackColor       =   &H006B4B01&
         BorderStyle     =   0  'None
         Height          =   1965
         Index           =   5
         Left            =   7350
         TabIndex        =   52
         Top             =   150
         Width           =   3885
         Begin VB.TextBox txtMiddleName 
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
            Left            =   1230
            MaxLength       =   25
            TabIndex        =   56
            Top             =   330
            Width           =   2565
         End
         Begin VB.TextBox txtRefNum 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   0
            Width           =   2565
         End
         Begin VB.TextBox txtPathologist 
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
            Left            =   1230
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   54
            Top             =   1650
            Width           =   2325
         End
         Begin VB.TextBox txtSpecimen 
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
            Left            =   1230
            MaxLength       =   20
            TabIndex        =   53
            Top             =   660
            Width           =   2565
         End
         Begin MSComCtl2.DTPicker dtpRequestDate 
            Height          =   285
            Left            =   1230
            TabIndex        =   57
            Top             =   1320
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   55836673
            UpDown          =   -1  'True
            CurrentDate     =   38406
         End
         Begin MSComCtl2.DTPicker dtpRequestTime 
            Height          =   285
            Left            =   2550
            TabIndex        =   58
            Top             =   1320
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:mm tt"
            Format          =   55836675
            UpDown          =   -1  'True
            CurrentDate     =   38406
         End
         Begin MSComCtl2.DTPicker dtpResultDate 
            Height          =   285
            Left            =   1230
            TabIndex        =   59
            Top             =   990
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   55836673
            UpDown          =   -1  'True
            CurrentDate     =   38406
         End
         Begin MSComCtl2.DTPicker dtpResultTime 
            Height          =   285
            Left            =   2550
            TabIndex        =   60
            Top             =   990
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:mm tt"
            Format          =   55836675
            UpDown          =   -1  'True
            CurrentDate     =   38406
         End
         Begin VB.Label lblText 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Middle Name:"
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
            Left            =   -30
            TabIndex        =   67
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblText 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FCE2C7&
            BackStyle       =   0  'Transparent
            Caption         =   "Ref. Number:"
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
            Left            =   -30
            TabIndex        =   66
            Top             =   60
            Width           =   1215
         End
         Begin VB.Label lblRequestDate 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FCE2C7&
            BackStyle       =   0  'Transparent
            Caption         =   "Request Date:"
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
            Left            =   -60
            TabIndex        =   65
            Top             =   1350
            Width           =   1230
         End
         Begin VB.Label lblResultdate 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FCE2C7&
            BackStyle       =   0  'Transparent
            Caption         =   "Result Date:"
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
            Left            =   -60
            TabIndex        =   64
            Top             =   1020
            Width           =   1230
         End
         Begin VB.Label lblText 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Pathologist:"
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
            Left            =   -30
            TabIndex        =   63
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblSpecimen 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FCE2C7&
            BackStyle       =   0  'Transparent
            Caption         =   "Specimen:"
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
            Left            =   -60
            TabIndex        =   62
            Top             =   690
            Width           =   1230
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "..."
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
            Height          =   315
            Left            =   3540
            TabIndex        =   61
            Top             =   1650
            Width           =   285
         End
      End
      Begin VB.Frame fraBody 
         BackColor       =   &H006B4B01&
         BorderStyle     =   0  'None
         Height          =   1965
         Index           =   4
         Left            =   3720
         TabIndex        =   33
         Top             =   150
         Width           =   3705
         Begin VB.TextBox txtFirstName 
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
            Left            =   1110
            MaxLength       =   25
            TabIndex        =   41
            Top             =   330
            Width           =   2565
         End
         Begin VB.TextBox txtMedtech 
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
            Left            =   1110
            MaxLength       =   30
            TabIndex        =   40
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox txtSpecimenID 
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
            Left            =   1110
            MaxLength       =   15
            TabIndex        =   39
            Top             =   0
            Width           =   2565
         End
         Begin VB.TextBox txtWard 
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
            Left            =   1110
            MaxLength       =   10
            TabIndex        =   38
            Top             =   990
            Width           =   1005
         End
         Begin VB.TextBox txtRoom 
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
            Left            =   2790
            MaxLength       =   10
            TabIndex        =   37
            Top             =   990
            Width           =   885
         End
         Begin VB.ListBox lstMedtech 
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
            Height          =   240
            Left            =   2490
            TabIndex        =   36
            Top             =   1320
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox txtClassID 
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
            Left            =   1110
            MaxLength       =   8
            TabIndex        =   35
            Top             =   1650
            Width           =   465
         End
         Begin VB.TextBox txtCivilStatus 
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
            Left            =   1110
            MaxLength       =   20
            TabIndex        =   34
            Top             =   660
            Width           =   2565
         End
         Begin VB.Label lblText 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FCE2C7&
            BackStyle       =   0  'Transparent
            Caption         =   "First Name:"
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
            Left            =   0
            TabIndex        =   51
            Top             =   330
            Width           =   1020
         End
         Begin VB.Label lblWard 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FCE2C7&
            BackStyle       =   0  'Transparent
            Caption         =   "Ward:"
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
            Left            =   0
            TabIndex        =   50
            Top             =   1020
            Width           =   1035
         End
         Begin VB.Label lblRoom 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FCE2C7&
            BackStyle       =   0  'Transparent
            Caption         =   "Room:"
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
            Left            =   2220
            TabIndex        =   49
            Top             =   1020
            Width           =   450
         End
         Begin VB.Label lblMedtech 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Med. Tech.:"
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
            Left            =   0
            TabIndex        =   48
            Top             =   1380
            Width           =   1050
         End
         Begin VB.Label lblSpecimenID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FCE2C7&
            BackStyle       =   0  'Transparent
            Caption         =   "Specimen ID:"
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
            Left            =   30
            TabIndex        =   47
            Top             =   60
            Width           =   990
         End
         Begin VB.Label lblText 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FCE2C7&
            BackStyle       =   0  'Transparent
            Caption         =   "Classification:"
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
            Left            =   0
            TabIndex        =   46
            Top             =   1680
            Width           =   1065
         End
         Begin VB.Label lblClassName 
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
            Left            =   1590
            TabIndex        =   45
            Top             =   1650
            Width           =   1815
         End
         Begin VB.Label lblCivilStatus 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FCE2C7&
            BackStyle       =   0  'Transparent
            Caption         =   "Civil Status:"
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
            Left            =   0
            TabIndex        =   44
            Top             =   660
            Width           =   1020
         End
         Begin VB.Label lblClassB 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "..."
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
            Height          =   315
            Left            =   3390
            TabIndex        =   43
            Top             =   1650
            Width           =   285
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "..."
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
            Height          =   315
            Left            =   3390
            TabIndex        =   42
            Top             =   1320
            Width           =   285
         End
      End
      Begin VB.Frame fraBody 
         BackColor       =   &H006B4B01&
         BorderStyle     =   0  'None
         Height          =   1965
         Index           =   6
         Left            =   30
         TabIndex        =   14
         Top             =   150
         Width           =   3615
         Begin VB.TextBox txtPatientID 
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
            Left            =   1050
            MaxLength       =   15
            TabIndex        =   22
            Top             =   0
            Width           =   2565
         End
         Begin VB.TextBox txtLastName 
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
            Left            =   1050
            MaxLength       =   25
            TabIndex        =   21
            Top             =   330
            Width           =   2565
         End
         Begin VB.TextBox txtPhysician 
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
            Left            =   1050
            MaxLength       =   25
            TabIndex        =   20
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox txtAge 
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
            Left            =   1050
            MaxLength       =   10
            TabIndex        =   19
            Top             =   990
            Width           =   975
         End
         Begin VB.TextBox txtSex 
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
            Left            =   2400
            MaxLength       =   6
            TabIndex        =   18
            Top             =   990
            Width           =   1215
         End
         Begin VB.TextBox txtProfileID 
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
            Left            =   1050
            MaxLength       =   8
            TabIndex        =   17
            Top             =   1650
            Width           =   375
         End
         Begin VB.ListBox lstPhysician 
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
            Height          =   240
            Left            =   2430
            TabIndex        =   16
            Top             =   1320
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox txtBirthDate 
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
            Left            =   1050
            MaxLength       =   30
            TabIndex        =   15
            Top             =   660
            Width           =   2565
         End
         Begin VB.Label lblText 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Last Name:"
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
            Left            =   0
            TabIndex        =   32
            Top             =   360
            Width           =   990
         End
         Begin VB.Label lblPhysician 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Physician:"
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
            Left            =   0
            TabIndex        =   31
            Top             =   1350
            Width           =   975
         End
         Begin VB.Label lblAge 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Age:"
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
            Left            =   0
            TabIndex        =   30
            Top             =   1020
            Width           =   975
         End
         Begin VB.Label lblSex 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FCE2C7&
            BackStyle       =   0  'Transparent
            Caption         =   "Sex:"
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
            Left            =   2070
            TabIndex        =   29
            Top             =   1020
            Width           =   330
         End
         Begin VB.Label lblText 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Index           =   4
            Left            =   30
            TabIndex        =   28
            Top             =   1680
            Width           =   960
         End
         Begin VB.Label lblProfileName 
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
            Left            =   1440
            TabIndex        =   27
            Top             =   1650
            Width           =   1905
         End
         Begin VB.Label lblPatientID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Patient ID:"
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
            Left            =   180
            TabIndex        =   26
            Top             =   30
            Width           =   795
         End
         Begin VB.Label lblBirthDate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Birth Date:"
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
            Left            =   0
            TabIndex        =   25
            Top             =   690
            Width           =   990
         End
         Begin VB.Label lblProfileidB 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "..."
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
            Height          =   315
            Left            =   3330
            TabIndex        =   24
            Top             =   1650
            Width           =   285
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "..."
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
            Height          =   315
            Left            =   3330
            TabIndex        =   23
            Top             =   1320
            Width           =   285
         End
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Position:"
      Height          =   195
      Left            =   1590
      TabIndex        =   71
      Top             =   9870
      Width           =   600
   End
   Begin VB.Label Label6 
      Caption         =   "Rack:"
      Height          =   225
      Left            =   90
      TabIndex        =   70
      Top             =   9840
      Width           =   495
   End
End
Attribute VB_Name = "frmWorkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rstResult As New ADODB.Recordset
Private lastRack As String
Private lastCass As String

Private Sub cboFilter_Click()
Dim myDate As String
myDate = Format(dtpDate.Value, "yyyymmdd")
myDate = myDate & "%"
Select Case cboFilter.Text
    Case "All"
        Set rstResult = myConnection.Execute("Select * from tblabrequest where messagedate like '" & myDate & "'")
    Case "Pending"
        Set rstResult = myConnection.Execute("Select * from tblabrequest where messagestatus = '0' and messagedate like '" & myDate & "'")
    Case "Sending Request"
        Set rstResult = myConnection.Execute("Select * from tblabrequest where messagestatus = '1' and messagedate like '" & myDate & "'")
    Case "Waiting"
        Set rstResult = myConnection.Execute("Select * from tblabrequest where messagestatus = '2' and messagedate like '" & myDate & "'")
    Case "With Result"
        Set rstResult = myConnection.Execute("Select * from tblabrequest where messagestatus = '3' and messagedate like '" & myDate & "'")
End Select
If rstResult.RecordCount > 0 Then
    Set msfgResult.DataSource = rstResult
    msfgResult.Visible = True
    setGrid
Else
    msfgResult.Visible = False
End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPos_Click()
Dim ctr  As Integer
Dim rstTemp As ADODB.Recordset
Set rstTemp = myConnection.Execute("Select * from tblabRequest where messagerack='" & txtRack.Text & "' and messagecasspos='" & txtCass.Text & "' and (messagestatus = '1' or messagestatus = '2')")
If rstTemp.RecordCount = 0 Then
    With msfgResult
        myConnection.Execute "update tblabRequest set MessageRack = '" & txtRack.Text & "', MessageCassPos = '" & txtCass.Text & "' where MessageId = " & .TextMatrix(.Row, 1)
        rstResult.Requery
        If rstResult.RecordCount > 0 Then
            Set msfgResult.Recordset = rstResult
            setGrid
            msfgResult.Visible = True
        Else
            msfgResult.Visible = False
        End If
    End With
    If MessageFormat = "STKS" Then
    ElseIf MessageFormat = "ASTM" Then
        If Val(txtCass.Text) < 9 Then
            txtCass.Text = Trim(Str(Val(txtCass.Text) + 1))
        Else
            txtCass.Text = 1
            txtRack.Text = Trim(Str(Val(txtRack.Text) + 1))
        End If
    End If
    lastRack = txtRack.Text
    lastCass = txtCass.Text
    If msfgResult.Row < msfgResult.Rows - 1 Then msfgResult.Row = msfgResult.Row + 1
Else
    If Trim(txtRack.Text) = "" And Trim(txtCass.Text) = "" Then
        With msfgResult
            myConnection.Execute "update tblabRequest set MessageRack = '" & txtRack.Text & "', MessageCassPos = '" & txtCass.Text & "' where MessageId = " & .TextMatrix(.Row, 1)
            rstResult.Requery
            If rstResult.RecordCount > 0 Then
                Set msfgResult.Recordset = rstResult
                setGrid
                msfgResult.Visible = True
            Else
                msfgResult.Visible = False
            End If
        End With
        If MessageFormat = "STKS" Then
        ElseIf MessageFormat = "ASTM" Then
            If Val(txtCass.Text) < 9 Then
                txtCass.Text = Trim(Str(Val(txtCass.Text) + 1))
            Else
                txtCass.Text = 1
                txtRack.Text = Trim(Str(Val(txtRack.Text) + 1))
            End If
        End If
        lastRack = txtRack.Text
        lastCass = txtCass.Text
        If msfgResult.Row < msfgResult.Rows - 1 Then msfgResult.Row = msfgResult.Row + 1
    Else
        MsgBox "Rack and Cassete Position Already Exist!", vbInformation + vbOKOnly, "Information"
    End If
End If
End Sub

Private Sub cmdRefresh_Click()
cboFilter_Click
End Sub

Private Sub cmdRemoveFromWorklist_Click()
With msfgResult
    myConnection.Execute "update tblabRequest set MessageStatus = 1, MessageTrans = 'C', MessageRack = '', MessageCassPos = '' where MessageId = " & .TextMatrix(.Row, 1)
    rstResult.Requery
    If rstResult.RecordCount > 0 Then
        Set msfgResult.Recordset = rstResult
        setGrid
        msfgResult.Visible = True
    Else
        msfgResult.Visible = False
    End If
End With
End Sub

Private Sub cmdSend_Click()
With msfgResult
    myConnection.Execute "update tblabRequest set MessageStatus = 1, MessageTrans = 'N' where MessageId = " & .TextMatrix(.Row, 1)
    rstResult.Requery
    If rstResult.RecordCount > 0 Then
        Set msfgResult.Recordset = rstResult
        setGrid
        msfgResult.Visible = True
    Else
        msfgResult.Visible = False
    End If
End With
End Sub


Private Sub dtpDate_Click()
cboFilter_Click
End Sub

Private Sub Form_Load()
'Set rstResult = myConnection.Execute("Select * from tblabrequest")
'Set msfgResult.DataSource = rstResult
'setGrid
If txtFilter.Visible Then txtFilter.SetFocus
fillCboFilter
txtRack.Text = ""
txtCass.Text = ""
dtpDate = Now
cmdRefresh_Click
End Sub

Private Sub setGrid()
With msfgResult
    .ColWidth(0) = 400
    .ColWidth(1) = 0
    .ColWidth(2) = 4000
    .ColWidth(3) = 0
    .ColWidth(4) = 0
    .ColWidth(5) = 0
    .ColWidth(6) = 500
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
    .ColWidth(10) = 2000
    .ColWidth(11) = 0
    .ColWidth(12) = 1000
    .ColWidth(13) = 1000
    .ColWidth(14) = 0
    .ColWidth(15) = 0
    .ColWidth(16) = 400
    .TextMatrix(0, 2) = "Name"
    .TextMatrix(0, 6) = "Sex"
    .TextMatrix(0, 7) = "DOB"
    .TextMatrix(0, 8) = "Date"
    .TextMatrix(0, 9) = "PIN"
    .TextMatrix(0, 10) = "SID"
    .TextMatrix(0, 12) = "Rack"
    .TextMatrix(0, 13) = "Cass"
    .TextMatrix(0, 15) = ""
    .TextMatrix(0, 16) = ""
End With
End Sub

Private Sub fillCboFilter()
With cboFilter
    .AddItem "All"
    .AddItem "Pending"
    .AddItem "Sending Request"
    .AddItem "Waiting"
    .AddItem "With Result"
    .ListIndex = 0
End With
End Sub

Private Sub msfgResult_Click()
    msfgResult_EnterCell
End Sub

Private Sub msfgResult_EnterCell()
Dim rstRequest As ADODB.Recordset
With msfgResult
    If .TextMatrix(.Row, 16) = 3 Or .TextMatrix(.Row, 16) = 1 Then
        cmdSend.Enabled = False
    Else
        cmdSend.Enabled = True
    End If
    Set rstRequest = myConnection.Execute("Select * from tbLabRequest where messageid = " & Trim(.TextMatrix(.Row, 1)))
    txtExams.Text = getExams(rstRequest.Fields("Messagedata").Value)
    If Trim(Len(.TextMatrix(.Row, 12))) > 0 Then
        txtRack.Text = .TextMatrix(.Row, 12)
    Else
        lastRack = txtRack.Text
    End If
    If Trim(Len(.TextMatrix(.Row, 13))) > 0 Then
        txtCass.Text = .TextMatrix(.Row, 13)
    Else
        lastCass = txtCass.Text
    End If
End With
End Sub

Private Sub msfgResult_KeyDown(KeyCode As Integer, Shift As Integer)
With msfgResult
    If KeyCode = vbKeyDelete And Shift = 1 Then
        If MsgBox("Are you sure you want to delete this item?", vbYesNo + vbQuestion + vbDefaultButton2, "Delete") = vbYes Then
            myConnection.Execute "Delete from TblabRequest where MessageId = " & .TextMatrix(.Row, 1)
            rstResult.Requery
            If rstResult.RecordCount > 0 Then
                Set msfgResult.Recordset = rstResult
                setGrid
                msfgResult.Visible = True
            Else
                msfgResult.Visible = False
            End If
        End If
    End If
End With
End Sub

Private Sub tmrRefresh_Timer()
cmdRefresh_Click
End Sub

Private Sub txtFilter_Change()
Dim tempStr As String
    tempStr = "[" & rstResult.Fields("MessageName").Name & "] like '%" & txtFilter.Text & "%'"
    If Len(Trim(txtFilter)) > 0 Then
        rstResult.Filter = tempStr
    Else
        rstResult.Filter = ""
    End If
    If rstResult.RecordCount = 0 Then
        msfgResult.Visible = False
        txtExams.Text = ""
    Else
        msfgResult.Visible = True
        Set msfgResult.Recordset = rstResult
    End If
    setGrid
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Function getExams(mess As String) As String
Dim arrTemp() As String
Dim arrTemp2() As String
Dim arrTemp3() As String
Dim ctr As Integer
Dim mExams As String
arrTemp = Split(mess, vbCr)
mExams = ""
For ctr = 0 To UBound(arrTemp) - 1
    arrTemp2 = Split(arrTemp(ctr), "|")
    If arrTemp2(0) = "OBR" Then
        arrTemp3 = Split(arrTemp2(4), "^")
        If Len(Trim(mExams)) = 0 Then
            mExams = arrTemp3(1)
        Else
            mExams = mExams & "," & arrTemp3(1)
        End If
    End If
Next ctr
getExams = mExams
End Function
