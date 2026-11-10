VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFormSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form Settings"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   7530
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7770
      TabIndex        =   72
      Top             =   1320
      Width           =   2295
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Form Settings"
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
         TabIndex        =   73
         Top             =   150
         Width           =   1335
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      ForeColor       =   &H80000008&
      Height          =   7005
      Index           =   0
      Left            =   0
      TabIndex        =   74
      Top             =   0
      Width           =   7515
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Settings"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4200
         TabIndex        =   33
         Top             =   6480
         Width           =   1485
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6225
         Left            =   180
         TabIndex        =   75
         Top             =   210
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   10980
         _Version        =   393216
         Style           =   1
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "Result Entry (Page 1)"
         TabPicture(0)   =   "frmFormSettings.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(1)=   "Frame3"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Result Entry (Page 2)"
         TabPicture(1)   =   "frmFormSettings.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame4"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Result Entry (Page 3)"
         TabPicture(2)   =   "frmFormSettings.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame1"
         Tab(2).Control(1)=   "Frame5"
         Tab(2).Control(2)=   "Frame6"
         Tab(2).ControlCount=   3
         Begin VB.Frame Frame6 
            Appearance      =   0  'Flat
            Caption         =   "Initial Field Position"
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
            Height          =   3225
            Left            =   -71370
            TabIndex        =   99
            Top             =   2910
            Width           =   3405
            Begin VB.ComboBox cboFpos 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   71
               Top             =   450
               Width           =   3015
            End
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            Caption         =   "Mandatory Fields"
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
            Height          =   3225
            Left            =   -74910
            TabIndex        =   98
            Top             =   2910
            Width           =   3405
            Begin VB.CheckBox chkMan06 
               Appearance      =   0  'Flat
               Caption         =   "Specimen"
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
               Height          =   210
               Left            =   270
               TabIndex        =   69
               Top             =   1890
               Width           =   2175
            End
            Begin VB.CheckBox chkMan05 
               Appearance      =   0  'Flat
               Caption         =   "Civil Status"
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
               Height          =   210
               Left            =   270
               TabIndex        =   68
               Top             =   1560
               Width           =   2175
            End
            Begin VB.CheckBox chkMan07 
               Appearance      =   0  'Flat
               Caption         =   "Routine/Stat"
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
               Height          =   210
               Left            =   270
               TabIndex        =   70
               Top             =   2220
               Width           =   2175
            End
            Begin VB.CheckBox chkMan01 
               Appearance      =   0  'Flat
               Caption         =   "Patient ID"
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
               Height          =   210
               Left            =   270
               TabIndex        =   64
               Top             =   360
               Width           =   2025
            End
            Begin VB.CheckBox chkMan02 
               Appearance      =   0  'Flat
               Caption         =   "Specimen ID"
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
               Height          =   210
               Left            =   270
               TabIndex        =   65
               Top             =   660
               Width           =   2025
            End
            Begin VB.CheckBox chkMan03 
               Appearance      =   0  'Flat
               Caption         =   "Ward"
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
               Height          =   210
               Left            =   270
               TabIndex        =   66
               Top             =   960
               Width           =   2025
            End
            Begin VB.CheckBox chkMan04 
               Appearance      =   0  'Flat
               Caption         =   "Room"
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
               Height          =   210
               Left            =   270
               TabIndex        =   67
               Top             =   1260
               Width           =   2025
            End
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   2505
            Left            =   -74910
            TabIndex        =   95
            Top             =   360
            Width           =   6975
            Begin VB.ComboBox Combo7 
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
               Height          =   330
               Left            =   1590
               Style           =   2  'Dropdown List
               TabIndex        =   63
               Top             =   1980
               Width           =   2505
            End
            Begin VB.CheckBox chkRoutineStat 
               Appearance      =   0  'Flat
               Caption         =   "Routine/Stat"
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
               Height          =   210
               Left            =   240
               TabIndex        =   62
               Top             =   2010
               Width           =   1485
            End
            Begin VB.ComboBox Combo6 
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
               Height          =   330
               Left            =   1590
               Style           =   2  'Dropdown List
               TabIndex        =   61
               Top             =   1620
               Width           =   2505
            End
            Begin VB.ComboBox Combo5 
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
               Height          =   330
               Left            =   1590
               Style           =   2  'Dropdown List
               TabIndex        =   59
               Top             =   1260
               Width           =   2505
            End
            Begin VB.ComboBox Combo4 
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
               Height          =   330
               Left            =   1590
               Style           =   2  'Dropdown List
               TabIndex        =   57
               Top             =   900
               Width           =   2505
            End
            Begin VB.ComboBox Combo3 
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
               Height          =   330
               Left            =   1590
               Style           =   2  'Dropdown List
               TabIndex        =   55
               Top             =   540
               Width           =   2505
            End
            Begin VB.ComboBox Combo2 
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
               Height          =   330
               Left            =   1590
               Style           =   2  'Dropdown List
               TabIndex        =   53
               Top             =   180
               Width           =   2505
            End
            Begin VB.CheckBox chkWard 
               Appearance      =   0  'Flat
               Caption         =   "Ward"
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
               Height          =   210
               Left            =   240
               TabIndex        =   52
               Top             =   210
               Width           =   1455
            End
            Begin VB.CheckBox chkRoom 
               Appearance      =   0  'Flat
               Caption         =   "Room"
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
               Height          =   210
               Left            =   240
               TabIndex        =   54
               Top             =   570
               Width           =   1455
            End
            Begin VB.CheckBox chkRemarks 
               Appearance      =   0  'Flat
               Caption         =   "Remarks"
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
               Height          =   210
               Left            =   240
               TabIndex        =   60
               Top             =   1650
               Width           =   1485
            End
            Begin VB.CheckBox chkSpecimen 
               Appearance      =   0  'Flat
               Caption         =   "Specimen"
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
               Height          =   210
               Left            =   240
               TabIndex        =   58
               Top             =   1290
               Width           =   1485
            End
            Begin VB.CheckBox chkCS 
               Appearance      =   0  'Flat
               Caption         =   "Civil Status"
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
               Height          =   210
               Left            =   240
               TabIndex        =   56
               Top             =   930
               Width           =   1485
            End
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   4905
            Left            =   150
            TabIndex        =   89
            Top             =   510
            Width           =   6795
            Begin VB.CheckBox Check32 
               Appearance      =   0  'Flat
               Caption         =   "Enable MT/RT Autofill"
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
               Left            =   4230
               TabIndex        =   100
               Top             =   1980
               Width           =   2295
            End
            Begin VB.CheckBox Check31 
               Appearance      =   0  'Flat
               Caption         =   "Save Last Used MT/RT"
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
               Left            =   2070
               TabIndex        =   46
               Top             =   1980
               Width           =   2295
            End
            Begin VB.CheckBox Check30 
               Appearance      =   0  'Flat
               Caption         =   "Use <NewPatientName> field in Result Entry "
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
               Left            =   180
               TabIndex        =   49
               Top             =   2910
               Width           =   3705
            End
            Begin VB.CheckBox Check29 
               Appearance      =   0  'Flat
               Caption         =   "Search OBI"
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
               Left            =   180
               TabIndex        =   48
               Top             =   2610
               Width           =   1635
            End
            Begin VB.CheckBox Check27 
               Appearance      =   0  'Flat
               Caption         =   "Swap Values of Patient ID and SpecimenID (New && Unedited Only)"
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
               Left            =   180
               TabIndex        =   47
               Top             =   2310
               Width           =   5235
            End
            Begin VB.CheckBox Check25 
               Appearance      =   0  'Flat
               Caption         =   "Auto Resize Form"
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
               Left            =   180
               TabIndex        =   45
               Top             =   2010
               Width           =   1635
            End
            Begin VB.CheckBox Check24 
               Appearance      =   0  'Flat
               Caption         =   "Update Search Result Automatically"
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
               Left            =   180
               TabIndex        =   44
               Top             =   1710
               Width           =   2985
            End
            Begin VB.CheckBox Check23 
               Appearance      =   0  'Flat
               Caption         =   "Auto Detect Sex for Normal Values Classification"
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
               Left            =   180
               TabIndex        =   43
               Top             =   1410
               Width           =   4185
            End
            Begin VB.CheckBox Check22 
               Appearance      =   0  'Flat
               Caption         =   "Show Print Copies"
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
               Left            =   180
               TabIndex        =   41
               Top             =   1110
               Width           =   1755
            End
            Begin VB.CheckBox Check21 
               Appearance      =   0  'Flat
               Caption         =   "Reverse SI  && Conventional Tag "
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
               Left            =   2070
               TabIndex        =   42
               Top             =   1110
               Width           =   2925
            End
            Begin VB.CheckBox Check20 
               Appearance      =   0  'Flat
               Caption         =   "Check ID from Patient Name"
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
               Left            =   2070
               TabIndex        =   40
               Top             =   810
               Width           =   2475
            End
            Begin VB.CheckBox Check19 
               Appearance      =   0  'Flat
               Caption         =   "Compute Age"
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
               Left            =   180
               TabIndex        =   39
               Top             =   810
               Width           =   2475
            End
            Begin VB.CheckBox Check17 
               Appearance      =   0  'Flat
               Caption         =   "Enable Patient ID Search"
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
               Left            =   2070
               TabIndex        =   38
               Top             =   510
               Width           =   2475
            End
            Begin VB.CheckBox Check16 
               Appearance      =   0  'Flat
               Caption         =   "Save Blank Results"
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
               Left            =   180
               TabIndex        =   37
               Top             =   510
               Width           =   1845
            End
            Begin VB.CheckBox Check15 
               Appearance      =   0  'Flat
               Caption         =   "Bypass Diff. Count Validation"
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
               Left            =   2070
               TabIndex        =   36
               Top             =   210
               Width           =   2565
            End
            Begin VB.TextBox txtPIDSuffix 
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
               Left            =   1740
               MaxLength       =   5
               TabIndex        =   51
               Top             =   3690
               Width           =   585
            End
            Begin VB.TextBox txtSIDSuffix 
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
               Left            =   1740
               MaxLength       =   5
               TabIndex        =   50
               Top             =   3300
               Width           =   585
            End
            Begin VB.CheckBox chkShowSI 
               Appearance      =   0  'Flat
               Caption         =   "Show SI Results"
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
               Left            =   180
               TabIndex        =   35
               Top             =   210
               Width           =   1665
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Patient ID Suffix:"
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
               Left            =   390
               TabIndex        =   91
               Top             =   3720
               Width           =   1320
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Specimen ID Suffix:"
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
               Left            =   150
               TabIndex        =   90
               Top             =   3330
               Width           =   1575
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   5835
            Left            =   -72270
            TabIndex        =   77
            Top             =   330
            Width           =   4335
            Begin VB.TextBox Text16 
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
               TabIndex        =   28
               Top             =   3780
               Width           =   3015
            End
            Begin VB.TextBox Text15 
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
               TabIndex        =   32
               Top             =   5100
               Width           =   3015
            End
            Begin VB.TextBox Text12 
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
               TabIndex        =   29
               Top             =   4110
               Width           =   3015
            End
            Begin VB.TextBox Text13 
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
               TabIndex        =   30
               Top             =   4440
               Width           =   3015
            End
            Begin VB.TextBox Text14 
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
               TabIndex        =   31
               Top             =   4770
               Width           =   3015
            End
            Begin VB.TextBox Text11 
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
               TabIndex        =   27
               Top             =   3450
               Width           =   3015
            End
            Begin VB.TextBox Text10 
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
               TabIndex        =   26
               Top             =   3120
               Width           =   3015
            End
            Begin VB.TextBox Text9 
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
               TabIndex        =   25
               Top             =   2790
               Width           =   3015
            End
            Begin VB.TextBox Text8 
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
               TabIndex        =   24
               Top             =   2460
               Width           =   3015
            End
            Begin VB.TextBox Text7 
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
               TabIndex        =   23
               Top             =   2130
               Width           =   3015
            End
            Begin VB.TextBox Text6 
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
               TabIndex        =   22
               Top             =   1800
               Width           =   3015
            End
            Begin VB.TextBox Text5 
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
               TabIndex        =   21
               Top             =   1470
               Width           =   3015
            End
            Begin VB.TextBox Text4 
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
               TabIndex        =   20
               Top             =   1140
               Width           =   3015
            End
            Begin VB.TextBox Text3 
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
               TabIndex        =   19
               Top             =   810
               Width           =   3015
            End
            Begin VB.TextBox Text2 
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
               TabIndex        =   18
               Top             =   480
               Width           =   3015
            End
            Begin VB.TextBox Text1 
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
               TabIndex        =   17
               Top             =   150
               Width           =   3015
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Remarks 2"
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
               Left            =   300
               TabIndex        =   97
               Top             =   3810
               Width           =   885
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Routine/Stat"
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
               Left            =   195
               TabIndex        =   96
               Top             =   5100
               Width           =   990
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Birth Date"
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
               Left            =   375
               TabIndex        =   94
               Top             =   4170
               Width           =   795
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Civil Status"
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
               Left            =   255
               TabIndex        =   93
               Top             =   4500
               Width           =   915
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Specimen"
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
               Left            =   345
               TabIndex        =   92
               Top             =   4800
               Width           =   825
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Remarks"
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
               Left            =   420
               TabIndex        =   88
               Top             =   3480
               Width           =   750
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Med. Tech."
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
               Left            =   270
               TabIndex        =   87
               Top             =   3180
               Width           =   900
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Physician"
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
               Left            =   390
               TabIndex        =   86
               Top             =   2850
               Width           =   780
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Request Date"
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
               Left            =   45
               TabIndex        =   85
               Top             =   2490
               Width           =   1095
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Result Date"
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
               Left            =   210
               TabIndex        =   84
               Top             =   2160
               Width           =   930
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Room"
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
               Left            =   660
               TabIndex        =   83
               Top             =   1830
               Width           =   480
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ward"
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
               Left            =   705
               TabIndex        =   82
               Top             =   1530
               Width           =   420
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sex"
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
               Left            =   840
               TabIndex        =   81
               Top             =   1140
               Width           =   300
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Age"
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
               Left            =   810
               TabIndex        =   80
               Top             =   840
               Width           =   330
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Specimen ID"
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
               Left            =   120
               TabIndex        =   79
               Top             =   540
               Width           =   1020
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Patient ID"
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
               Left            =   360
               TabIndex        =   78
               Top             =   210
               Width           =   765
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   5835
            Left            =   -74940
            TabIndex        =   76
            Top             =   330
            Width           =   2655
            Begin VB.CheckBox Check28 
               Appearance      =   0  'Flat
               Caption         =   "Enable Remarks 2"
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
               Height          =   210
               Left            =   90
               TabIndex        =   11
               Top             =   3810
               Width           =   2175
            End
            Begin VB.CheckBox Check26 
               Appearance      =   0  'Flat
               Caption         =   "Enable Routine/Stat"
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
               Height          =   210
               Left            =   90
               TabIndex        =   16
               Top             =   5460
               Width           =   2175
            End
            Begin VB.CheckBox Check18 
               Appearance      =   0  'Flat
               Caption         =   "Enable Middle Name"
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
               Height          =   210
               Left            =   90
               TabIndex        =   15
               Top             =   5130
               Width           =   2175
            End
            Begin VB.CheckBox Check12 
               Appearance      =   0  'Flat
               Caption         =   "Enable BirthDate"
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
               Height          =   210
               Left            =   90
               TabIndex        =   12
               Top             =   4170
               Width           =   2175
            End
            Begin VB.CheckBox Check13 
               Appearance      =   0  'Flat
               Caption         =   "Enable Civil Status"
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
               Height          =   210
               Left            =   90
               TabIndex        =   13
               Top             =   4500
               Width           =   2175
            End
            Begin VB.CheckBox Check14 
               Appearance      =   0  'Flat
               Caption         =   "Enable Specimen"
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
               Height          =   210
               Left            =   90
               TabIndex        =   14
               Top             =   4830
               Width           =   2175
            End
            Begin VB.CheckBox Check11 
               Appearance      =   0  'Flat
               Caption         =   "Enable Remarks"
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
               Height          =   210
               Left            =   90
               TabIndex        =   10
               Top             =   3480
               Width           =   2175
            End
            Begin VB.CheckBox Check10 
               Appearance      =   0  'Flat
               Caption         =   "Enable Med. Tech."
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
               Height          =   210
               Left            =   90
               TabIndex        =   9
               Top             =   3150
               Width           =   2175
            End
            Begin VB.CheckBox Check9 
               Appearance      =   0  'Flat
               Caption         =   "Enable Physician"
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
               Height          =   210
               Left            =   90
               TabIndex        =   8
               Top             =   2820
               Width           =   2175
            End
            Begin VB.CheckBox Check8 
               Appearance      =   0  'Flat
               Caption         =   "Enable Request Date"
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
               Height          =   210
               Left            =   90
               TabIndex        =   7
               Top             =   2490
               Width           =   2175
            End
            Begin VB.CheckBox Check7 
               Appearance      =   0  'Flat
               Caption         =   "Enable Result Date"
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
               Height          =   210
               Left            =   90
               TabIndex        =   6
               Top             =   2160
               Width           =   2175
            End
            Begin VB.CheckBox Check6 
               Appearance      =   0  'Flat
               Caption         =   "Enable Room"
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
               Height          =   210
               Left            =   90
               TabIndex        =   5
               Top             =   1830
               Width           =   2025
            End
            Begin VB.CheckBox Check5 
               Appearance      =   0  'Flat
               Caption         =   "Enable Ward"
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
               Height          =   210
               Left            =   90
               TabIndex        =   4
               Top             =   1530
               Width           =   2025
            End
            Begin VB.CheckBox Check4 
               Appearance      =   0  'Flat
               Caption         =   "Enable Sex"
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
               Height          =   210
               Left            =   90
               TabIndex        =   3
               Top             =   1200
               Width           =   2025
            End
            Begin VB.CheckBox Check3 
               Appearance      =   0  'Flat
               Caption         =   "Enable Age"
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
               Height          =   210
               Left            =   90
               TabIndex        =   2
               Top             =   870
               Width           =   2025
            End
            Begin VB.CheckBox Check2 
               Appearance      =   0  'Flat
               Caption         =   "Enable Specimen ID"
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
               Height          =   210
               Left            =   90
               TabIndex        =   1
               Top             =   540
               Width           =   2025
            End
            Begin VB.CheckBox Check1 
               Appearance      =   0  'Flat
               Caption         =   "Enable Patient ID"
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
               Height          =   210
               Left            =   90
               TabIndex        =   0
               Top             =   240
               Width           =   2025
            End
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
         Height          =   435
         Left            =   5790
         TabIndex        =   34
         Top             =   6480
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmFormSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkCS_Click()
    Combo4.Enabled = chkCS.Value
End Sub

Private Sub chkRemarks_Click()
    Combo6.Enabled = chkRemarks.Value
End Sub

Private Sub chkRoom_Click()
    Combo3.Enabled = chkRoom.Value
End Sub

Private Sub chkRoutineStat_Click()
    Combo7.Enabled = chkRoutineStat.Value
End Sub

Private Sub chkSpecimen_Click()
    Combo5.Enabled = chkSpecimen.Value
End Sub

Private Sub chkWard_Click()
    Combo2.Enabled = chkWard.Value
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    SaveSetting "HostForms_SA", "Settings", "PatientID", Trim(str(Check1.Value))
    SaveSetting "HostForms_SA", "Settings", "MedTech", Trim(str(Check10.Value))
    SaveSetting "HostForms_SA", "Settings", "Remarks", Trim(str(Check11.Value))
    SaveSetting "HostForms_SA", "Settings", "Remarks2", Trim(str(Check28.Value))
    SaveSetting "HostForms_SA", "Settings", "AccNum", Trim(str(Check2.Value))
    SaveSetting "HostForms_SA", "Settings", "Age", Trim(str(Check3.Value))
    SaveSetting "HostForms_SA", "Settings", "Sex", Trim(str(Check4.Value))
    SaveSetting "HostForms_SA", "Settings", "Ward", Trim(str(Check5.Value))
    SaveSetting "HostForms_SA", "Settings", "Room", Trim(str(Check6.Value))
    SaveSetting "HostForms_SA", "Settings", "ResultDate", Trim(str(Check7.Value))
    SaveSetting "HostForms_SA", "Settings", "RequestDate", Trim(str(Check8.Value))
    SaveSetting "HostForms_SA", "Settings", "Physician", Trim(str(Check9.Value))
    SaveSetting "HostForms_SA", "Settings", "BirthDate", Trim(str(Check12.Value))
    SaveSetting "HostForms_SA", "Settings", "CivilStatus", Trim(str(Check13.Value))
    SaveSetting "HostForms_SA", "Settings", "Specimen", Trim(str(Check14.Value))
    SaveSetting "HostForms_SA", "Settings", "ByPassDiff", Trim(str(Check15.Value))
    SaveSetting "HostForms_SA", "Settings", "SaveBlankResult", Trim(str(Check16.Value))
    SaveSetting "HostForms_SA", "Settings", "SearchID", Trim(str(Check17.Value))
    SaveSetting "HostForms_SA", "Settings", "MiddleName", Trim(str(Check18.Value))
    SaveSetting "HostForms_SA", "Settings", "ComputeAge", Trim(str(Check19.Value))
    SaveSetting "HostForms_SA", "Settings", "CheckID", Trim(str(Check20.Value))
    SaveSetting "HostForms_SA", "Settings", "ReverseSI", Trim(str(Check21.Value))
    SaveSetting "HostForms_SA", "Settings", "ShowPrintCopy", Trim(str(Check22.Value))
    SaveSetting "HostForms_SA", "Settings", "DetectNV", Trim(str(Check23.Value))
    SaveSetting "HostForms_SA", "Settings", "UpdateSearchResult", Trim(str(Check24.Value))
    SaveSetting "HostForms_SA", "Settings", "AutoResize", Trim(str(Check25.Value))
    SaveSetting "HostForms_SA", "Settings", "RoutineStat", Trim(str(Check26.Value))
    SaveSetting "HostForms_SA", "Settings", "SwapID", Trim(str(Check27.Value))
    SaveSetting "HostForms_SA", "Settings", "Search OBI", Trim(str(Check29.Value))
    SaveSetting "HostForms_SA", "Settings", "UseNewName", Trim(str(Check30.Value))
    SaveSetting "HostForms_SA", "Settings", "SaveDefaultMT", Trim(str(Check31.Value))
    SaveSetting "HostForms_SA", "Settings", "AutofillMT", Trim(str(Check32.Value))
    
    SaveSetting "HostForms_SA", "LabelSettings", "PatientID", Text1.Text
    SaveSetting "HostForms_SA", "LabelSettings", "SpecimenID", Text2.Text
    SaveSetting "HostForms_SA", "LabelSettings", "Age", Text3.Text
    SaveSetting "HostForms_SA", "LabelSettings", "Sex", Text4.Text
    SaveSetting "HostForms_SA", "LabelSettings", "Ward", Text5.Text
    SaveSetting "HostForms_SA", "LabelSettings", "Room", Text6.Text
    SaveSetting "HostForms_SA", "LabelSettings", "Resultdate", Text7.Text
    SaveSetting "HostForms_SA", "LabelSettings", "RequestDate", Text8.Text
    SaveSetting "HostForms_SA", "LabelSettings", "Physician", Text9.Text
    SaveSetting "HostForms_SA", "LabelSettings", "Medtech", Text10.Text
    SaveSetting "HostForms_SA", "LabelSettings", "Remarks", Text11.Text
    SaveSetting "HostForms_SA", "LabelSettings", "Remarks2", Text16.Text
    SaveSetting "HostForms_SA", "LabelSettings", "BirthDate", Text12.Text
    SaveSetting "HostForms_SA", "LabelSettings", "CivilStatus", Text13.Text
    SaveSetting "HostForms_SA", "LabelSettings", "Specimen", Text14.Text
    SaveSetting "HostForms_SA", "LabelSettings", "RoutineStat", Text15.Text
    
    SaveSetting "HostForms_SA", "Settings", "ShowSI", Trim(str(chkShowSI.Value))
    SaveSetting "HostForms_SA", "Settings", "SIDSuffix", txtSIDSuffix.Text
    SaveSetting "HostForms_SA", "Settings", "PIDSuffix", txtPIDSuffix.Text

    SaveSetting "HostForms_SA", "SelList", "Ward", Combo2.Text
    SaveSetting "HostForms_SA", "SelList", "Room", Combo3.Text
    SaveSetting "HostForms_SA", "SelList", "Civil Status", Combo4.Text
    SaveSetting "HostForms_SA", "SelList", "Specimen", Combo5.Text
    SaveSetting "HostForms_SA", "SelList", "Remarks", Combo6.Text
    SaveSetting "HostForms_SA", "SelList", "RoutineStat", Combo7.Text

    SaveSetting "HostForms_SA", "SelList", "cWard", chkWard.Value
    SaveSetting "HostForms_SA", "SelList", "cRoom", chkRoom.Value
    SaveSetting "HostForms_SA", "SelList", "cCivil Status", chkCS.Value
    SaveSetting "HostForms_SA", "SelList", "cSpecimen", chkSpecimen.Value
    SaveSetting "HostForms_SA", "SelList", "cRemarks", chkRemarks.Value
    SaveSetting "HostForms_SA", "SelList", "cRoutineStat", chkRoutineStat.Value
    
    SaveSetting "HostForms_SA", "Settings", "ManFld01", Trim(str(chkMan01.Value))
    SaveSetting "HostForms_SA", "Settings", "ManFld02", Trim(str(chkMan02.Value))
    SaveSetting "HostForms_SA", "Settings", "ManFld03", Trim(str(chkMan03.Value))
    SaveSetting "HostForms_SA", "Settings", "ManFld04", Trim(str(chkMan04.Value))
    SaveSetting "HostForms_SA", "Settings", "ManFld05", Trim(str(chkMan05.Value))
    SaveSetting "HostForms_SA", "Settings", "ManFld06", Trim(str(chkMan06.Value))
    SaveSetting "HostForms_SA", "Settings", "ManFld07", Trim(str(chkMan07.Value))

    SaveSetting "HostForms_SA", "Settings", "DefaultFieldPos", cboFpos.Text
        
    MsgBox "Settings Saved.", vbOKOnly + vbInformation, "Form Settings"
End Sub

Private Sub Form_Load()
setFormColors Me
    On Error Resume Next
    Check1.Value = Val(GetSetting("HostForms_SA", "Settings", "PatientID", "1"))
    Check2.Value = Val(GetSetting("HostForms_SA", "Settings", "AccNum", "1"))
    Check3.Value = Val(GetSetting("HostForms_SA", "Settings", "Age", "1"))
    Check4.Value = Val(GetSetting("HostForms_SA", "Settings", "Sex", "1"))
    Check5.Value = Val(GetSetting("HostForms_SA", "Settings", "Ward", "1"))
    Check6.Value = Val(GetSetting("HostForms_SA", "Settings", "Room", "1"))
    Check7.Value = Val(GetSetting("HostForms_SA", "Settings", "ResultDate", "1"))
    Check8.Value = Val(GetSetting("HostForms_SA", "Settings", "RequestDate", "1"))
    Check9.Value = Val(GetSetting("HostForms_SA", "Settings", "Physician", "1"))
    Check10.Value = Val(GetSetting("HostForms_SA", "Settings", "MedTech", "1"))
    Check11.Value = Val(GetSetting("HostForms_SA", "Settings", "Remarks", "1"))
    Check28.Value = Val(GetSetting("HostForms_SA", "Settings", "Remarks2", "0"))
    Check12.Value = Val(GetSetting("HostForms_SA", "Settings", "BirthDate", "1"))
    Check13.Value = Val(GetSetting("HostForms_SA", "Settings", "CivilStatus", "1"))
    Check14.Value = Val(GetSetting("HostForms_SA", "Settings", "Specimen", "1"))
    Check15.Value = Val(GetSetting("HostForms_SA", "Settings", "ByPassDiff", "0"))
    Check16.Value = Val(GetSetting("HostForms_SA", "Settings", "SaveBlankResult", "1"))
    Check17.Value = Val(GetSetting("HostForms_SA", "Settings", "SearchID", "0"))
    Check18.Value = Val(GetSetting("HostForms_SA", "Settings", "MiddleName", "1"))
    Check19.Value = Val(GetSetting("HostForms_SA", "Settings", "ComputeAge", "1"))
    Check20.Value = Val(GetSetting("HostForms_SA", "Settings", "CheckID", "0"))
    Check21.Value = Val(GetSetting("HostForms_SA", "Settings", "ReverseSI", "0"))
    Check22.Value = Val(GetSetting("HostForms_SA", "Settings", "ShowPrintCopy", "0"))
    Check23.Value = Val(GetSetting("HostForms_SA", "Settings", "DetectNV", "0"))
    Check24.Value = Val(GetSetting("HostForms_SA", "Settings", "UpdateSearchResult", "0"))
    Check25.Value = Val(GetSetting("HostForms_SA", "Settings", "AutoResize", "0"))
    Check26.Value = Val(GetSetting("HostForms_SA", "Settings", "RoutineStat", "1"))
    Check27.Value = Val(GetSetting("HostForms_SA", "Settings", "SwapID", "0"))
    Check29.Value = Val(GetSetting("HostForms_SA", "Settings", "Search OBI", "0"))
    Check30.Value = Val(GetSetting("HostForms_SA", "Settings", "UseNewName", "0"))
    Check31.Value = Val(GetSetting("HostForms_SA", "Settings", "SaveDefaultMT", "0"))
    Check32.Value = Val(GetSetting("HostForms_SA", "Settings", "AutofillMT", "1"))
    
    chkShowSI.Value = Val(GetSetting("HostForms_SA", "Settings", "ShowSI", "0"))
    Text1.Text = GetSetting("HostForms_SA", "LabelSettings", "PatientID", "Patient ID")
    Text2.Text = GetSetting("HostForms_SA", "LabelSettings", "SpecimenID", "Specimen ID")
    Text3.Text = GetSetting("HostForms_SA", "LabelSettings", "Age", "Age")
    Text4.Text = GetSetting("HostForms_SA", "LabelSettings", "Sex", "Sex")
    Text5.Text = GetSetting("HostForms_SA", "LabelSettings", "Ward", "Ward")
    Text6.Text = GetSetting("HostForms_SA", "LabelSettings", "Room", "Room")
    Text7.Text = GetSetting("HostForms_SA", "LabelSettings", "Resultdate", "Result Date")
    Text8.Text = GetSetting("HostForms_SA", "LabelSettings", "RequestDate", "Request Date")
    Text9.Text = GetSetting("HostForms_SA", "LabelSettings", "Physician", "Physician")
    Text10.Text = GetSetting("HostForms_SA", "LabelSettings", "Medtech", "Med. Tech.")
    Text11.Text = GetSetting("HostForms_SA", "LabelSettings", "Remarks", "Remarks")
    Text16.Text = GetSetting("HostForms_SA", "LabelSettings", "Remarks2", "Remarks2")
    Text12.Text = GetSetting("HostForms_SA", "LabelSettings", "BirthDate", "Birth Date")
    Text13.Text = GetSetting("HostForms_SA", "LabelSettings", "CivilStatus", "Civil Status")
    Text14.Text = GetSetting("HostForms_SA", "LabelSettings", "Specimen", "Specimen")
    Text15.Text = GetSetting("HostForms_SA", "LabelSettings", "RoutineStat", "Routine/Stat")
    
    txtSIDSuffix.Text = GetSetting("HostForms_SA", "Settings", "SIDSuffix", "")
    txtPIDSuffix.Text = GetSetting("HostForms_SA", "Settings", "PIDSuffix", "")

    chkWard.Value = GetSetting("HostForms_SA", "SelList", "cWard", chkWard.Value)
    chkRoom.Value = GetSetting("HostForms_SA", "SelList", "cRoom", chkRoom.Value)
    chkCS.Value = GetSetting("HostForms_SA", "SelList", "cCivil Status", chkCS.Value)
    chkSpecimen.Value = GetSetting("HostForms_SA", "SelList", "cSpecimen", chkSpecimen.Value)
    chkRemarks.Value = GetSetting("HostForms_SA", "SelList", "cRemarks", chkRemarks.Value)
    chkRoutineStat.Value = GetSetting("HostForms_SA", "SelList", "cRoutineStat", chkRoutineStat.Value)
    
    fillCboSelection
    Combo2.ListIndex = findCboIndex(Combo2, GetSetting("HostForms_SA", "SelList", "Ward", "Wards"))
    Combo3.ListIndex = findCboIndex(Combo3, GetSetting("HostForms_SA", "SelList", "Room", "Rooms"))
    Combo4.ListIndex = findCboIndex(Combo4, GetSetting("HostForms_SA", "SelList", "Civil Status", "Civil Status"))
    Combo5.ListIndex = findCboIndex(Combo5, GetSetting("HostForms_SA", "SelList", "Specimen", "Specimen"))
    Combo6.ListIndex = findCboIndex(Combo6, GetSetting("HostForms_SA", "SelList", "Remarks", "Remarks"))
    Combo7.ListIndex = findCboIndex(Combo7, GetSetting("HostForms_SA", "SelList", "RoutineStat", "Routine/Stat"))
    
    chkMan01.Value = Val(GetSetting("HostForms_SA", "Settings", "ManFld01", "0"))
    chkMan02.Value = Val(GetSetting("HostForms_SA", "Settings", "ManFld02", "0"))
    chkMan03.Value = Val(GetSetting("HostForms_SA", "Settings", "ManFld03", "0"))
    chkMan04.Value = Val(GetSetting("HostForms_SA", "Settings", "ManFld04", "0"))
    chkMan05.Value = Val(GetSetting("HostForms_SA", "Settings", "ManFld05", "0"))
    chkMan06.Value = Val(GetSetting("HostForms_SA", "Settings", "ManFld06", "0"))
    chkMan07.Value = Val(GetSetting("HostForms_SA", "Settings", "ManFld07", "0"))
    
    With cboFpos
        .Clear
        .AddItem "Patient ID"
        .AddItem "Sample ID"
        .AddItem "Last Name"
        .AddItem "First Name"
    End With
    cboFpos.ListIndex = findCboIndex(cboFpos, GetSetting("HostForms_SA", "Settings", "DefaultFieldPos", "Patient ID"))
End Sub

Private Sub fillCboSelection()
Dim rstTypes As ADODB.Recordset
Set rstTypes = myConnection.Execute("Select * from tbSelectionType where status = '1'")
If rstTypes.RecordCount > 0 Then
    rstTypes.MoveFirst
    Do While Not rstTypes.EOF
        Combo2.AddItem rstTypes.Fields("SelType").Value
        Combo3.AddItem rstTypes.Fields("SelType").Value
        Combo4.AddItem rstTypes.Fields("SelType").Value
        Combo5.AddItem rstTypes.Fields("SelType").Value
        Combo6.AddItem rstTypes.Fields("SelType").Value
        Combo7.AddItem rstTypes.Fields("SelType").Value
        rstTypes.MoveNext
    Loop
End If
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
End With
End Function

