VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Settings"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7515
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      Caption         =   "Result Editing Method"
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   450
      TabIndex        =   96
      Top             =   5520
      Width           =   2985
      Begin VB.OptionButton optMethodA 
         Appearance      =   0  'Flat
         Caption         =   "Method A (Version 2.0.22)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   98
         Top             =   300
         Width           =   2565
      End
      Begin VB.OptionButton optMethodB 
         Appearance      =   0  'Flat
         Caption         =   "Method B (Newer Version)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   97
         Top             =   630
         Width           =   2565
      End
   End
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
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
      Height          =   495
      Left            =   8190
      TabIndex        =   21
      Top             =   1320
      Width           =   2685
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Program Settings"
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
         TabIndex        =   22
         Top             =   150
         Width           =   1695
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
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
      Height          =   4995
      Index           =   0
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   7515
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Settings"
         Height          =   435
         Left            =   4320
         TabIndex        =   19
         Top             =   4440
         Width           =   1425
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   435
         Left            =   5790
         TabIndex        =   20
         Top             =   4440
         Width           =   1545
      End
      Begin TabDlg.SSTab sstSettings 
         Height          =   4125
         Left            =   90
         TabIndex        =   24
         Top             =   210
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   7276
         _Version        =   393216
         Style           =   1
         Tabs            =   6
         TabsPerRow      =   6
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "frmSettings2.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtDetails"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label10"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label26"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label34"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Frame6"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtDiffCountTotal"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Frame2"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtProfileID"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Frame4"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "cmdSetPlugins"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txtClass"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "txtDiffFormat"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).ControlCount=   13
         TabCaption(1)   =   "Instrument Plugins"
         TabPicture(1)   =   "frmSettings2.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame8"
         Tab(1).Control(1)=   "cmdTestPlugin"
         Tab(1).Control(2)=   "cmdPluginSettings"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "DT Plugins"
         TabPicture(2)   =   "frmSettings2.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame3"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "VS Plugins"
         TabPicture(3)   =   "frmSettings2.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label27"
         Tab(3).Control(1)=   "Label28"
         Tab(3).Control(2)=   "Frame7"
         Tab(3).Control(3)=   "chkExport"
         Tab(3).Control(4)=   "txtLocation"
         Tab(3).Control(5)=   "cboExport"
         Tab(3).Control(6)=   "cdgExportLocation"
         Tab(3).Control(7)=   "cmdBrowse"
         Tab(3).ControlCount=   8
         TabCaption(4)   =   "Form Colors"
         TabPicture(4)   =   "frmSettings2.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame10"
         Tab(4).Control(1)=   "Frame1"
         Tab(4).Control(2)=   "Frame9"
         Tab(4).Control(3)=   "cdgColor"
         Tab(4).ControlCount=   4
         TabCaption(5)   =   "Others"
         TabPicture(5)   =   "frmSettings2.frx":008C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Label32"
         Tab(5).Control(1)=   "Label33"
         Tab(5).Control(2)=   "txtMessage"
         Tab(5).Control(3)=   "txtSamples"
         Tab(5).Control(4)=   "chkWarning"
         Tab(5).Control(5)=   "cmdSetting"
         Tab(5).Control(6)=   "Frame11"
         Tab(5).ControlCount=   7
         Begin VB.TextBox txtDiffFormat 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1470
            TabIndex        =   1
            Top             =   810
            Width           =   735
         End
         Begin VB.Frame Frame11 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   1965
            Left            =   -71850
            TabIndex        =   110
            Top             =   2010
            Width           =   3945
            Begin VB.CommandButton cmdTaskMAn 
               Caption         =   "Task Manager"
               Enabled         =   0   'False
               Height          =   405
               Left            =   1200
               TabIndex        =   116
               Top             =   1500
               Width           =   1305
            End
            Begin VB.CheckBox chkTaskM 
               Appearance      =   0  'Flat
               Caption         =   "Disable Task Manager"
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   120
               TabIndex        =   115
               Top             =   780
               Width           =   2355
            End
            Begin VB.CheckBox Check1 
               Appearance      =   0  'Flat
               Caption         =   "Disable Folder Options"
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   120
               TabIndex        =   114
               Top             =   1050
               Width           =   2445
            End
            Begin VB.CheckBox chkUsbStor 
               Appearance      =   0  'Flat
               Caption         =   "Disable USB Mass Storage"
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   120
               TabIndex        =   113
               Top             =   510
               Width           =   2355
            End
            Begin VB.CommandButton cmdRegistry 
               Caption         =   " Registry Editor"
               Enabled         =   0   'False
               Height          =   405
               Left            =   2550
               TabIndex        =   112
               Top             =   1500
               Width           =   1305
            End
            Begin VB.CheckBox chkRegistry 
               Appearance      =   0  'Flat
               Caption         =   "Disable Registry"
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   120
               TabIndex        =   111
               Top             =   240
               Width           =   1755
            End
         End
         Begin VB.CommandButton cmdSetting 
            Caption         =   "External Request Database Settings"
            Height          =   465
            Left            =   -74850
            TabIndex        =   109
            Top             =   3450
            Width           =   2805
         End
         Begin VB.CheckBox chkWarning 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Show Warning"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   -74880
            TabIndex        =   106
            Top             =   450
            Width           =   1545
         End
         Begin VB.TextBox txtSamples 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   -73560
            TabIndex        =   104
            Top             =   2100
            Width           =   555
         End
         Begin VB.CommandButton cmdPluginSettings 
            Caption         =   "Plugin Settings"
            Height          =   345
            Left            =   -69750
            TabIndex        =   103
            Top             =   3720
            Width           =   1875
         End
         Begin VB.CommandButton cmdBrowse 
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
            Height          =   315
            Left            =   -69480
            TabIndex        =   88
            Top             =   960
            Width           =   285
         End
         Begin MSComDlg.CommonDialog cdgExportLocation 
            Left            =   -68310
            Top             =   660
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.ComboBox cboExport 
            Height          =   330
            ItemData        =   "frmSettings2.frx":00A8
            Left            =   -74130
            List            =   "frmSettings2.frx":00AA
            Style           =   2  'Dropdown List
            TabIndex        =   87
            Top             =   1320
            Width           =   4665
         End
         Begin VB.TextBox txtLocation 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   -74130
            TabIndex        =   84
            Top             =   960
            Width           =   4635
         End
         Begin VB.CheckBox chkExport 
            Appearance      =   0  'Flat
            Caption         =   "Export"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   -74880
            TabIndex        =   83
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtClass 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1470
            TabIndex        =   3
            Top             =   1470
            Width           =   735
         End
         Begin VB.CommandButton cmdTestPlugin 
            Caption         =   "Test Plugins"
            Height          =   345
            Left            =   -74940
            TabIndex        =   81
            Top             =   3720
            Width           =   1875
         End
         Begin MSComDlg.CommonDialog cdgColor 
            Left            =   -74940
            Top             =   3750
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Frame Frame9 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   1425
            Left            =   -71280
            TabIndex        =   76
            Top             =   690
            Width           =   3465
            Begin VB.PictureBox picPrinted 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               ForeColor       =   &H80000008&
               Height          =   435
               Left            =   1140
               ScaleHeight     =   405
               ScaleWidth      =   2235
               TabIndex        =   78
               Top             =   720
               Width           =   2265
            End
            Begin VB.PictureBox picSaved 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               FillColor       =   &H0080FFFF&
               ForeColor       =   &H80000008&
               Height          =   435
               Left            =   1140
               ScaleHeight     =   405
               ScaleWidth      =   2235
               TabIndex        =   77
               Top             =   270
               Width           =   2265
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Printed Result"
               Height          =   210
               Left            =   60
               TabIndex        =   80
               Top             =   780
               Width           =   990
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Saved Result"
               Height          =   210
               Left            =   90
               TabIndex        =   79
               Top             =   360
               Width           =   960
            End
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   3375
            Left            =   -74970
            TabIndex        =   65
            Top             =   690
            Width           =   3645
            Begin VB.PictureBox picTitleBar 
               Appearance      =   0  'Flat
               BackColor       =   &H00A66B2B&
               ForeColor       =   &H80000008&
               Height          =   435
               Left            =   1140
               ScaleHeight     =   405
               ScaleWidth      =   2325
               TabIndex        =   70
               Top             =   1710
               Visible         =   0   'False
               Width           =   2355
            End
            Begin VB.PictureBox picTitleText 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H80000008&
               Height          =   435
               Left            =   1140
               ScaleHeight     =   405
               ScaleWidth      =   2325
               TabIndex        =   69
               Top             =   2190
               Visible         =   0   'False
               Width           =   2355
            End
            Begin VB.PictureBox picBodyText 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H80000008&
               Height          =   435
               Left            =   1170
               ScaleHeight     =   405
               ScaleWidth      =   2325
               TabIndex        =   68
               Top             =   1170
               Width           =   2355
            End
            Begin VB.PictureBox picBodyFrame 
               Appearance      =   0  'Flat
               BackColor       =   &H006B4B01&
               ForeColor       =   &H80000008&
               Height          =   435
               Left            =   1170
               ScaleHeight     =   405
               ScaleWidth      =   2325
               TabIndex        =   67
               Top             =   690
               Width           =   2355
            End
            Begin VB.ComboBox cboScheme 
               Height          =   330
               Left            =   1170
               Style           =   2  'Dropdown List
               TabIndex        =   66
               Top             =   210
               Width           =   2325
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Title Bar"
               Height          =   210
               Left            =   450
               TabIndex        =   75
               Top             =   1830
               Visible         =   0   'False
               Width           =   585
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Title Bar Text"
               Height          =   210
               Left            =   120
               TabIndex        =   74
               Top             =   2310
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Body Text"
               Height          =   210
               Left            =   360
               TabIndex        =   73
               Top             =   1260
               Width           =   735
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Body Frame"
               Height          =   210
               Left            =   240
               TabIndex        =   72
               Top             =   810
               Width           =   870
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Color Scheme"
               Height          =   210
               Left            =   90
               TabIndex        =   71
               Top             =   270
               Width           =   1005
            End
         End
         Begin VB.Frame Frame8 
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
            ForeColor       =   &H80000008&
            Height          =   2805
            Left            =   -74970
            TabIndex        =   44
            Top             =   840
            Width           =   7095
            Begin VB.FileListBox fileList 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   2130
               Left            =   120
               Pattern         =   "HOST_*.dll"
               TabIndex        =   51
               Top             =   480
               Width           =   1785
            End
            Begin VB.CheckBox chkPlugin 
               Appearance      =   0  'Flat
               Caption         =   "Enable Plugin"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   90
               TabIndex        =   50
               Top             =   180
               Width           =   1785
            End
            Begin VB.TextBox txtInstrumentID 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3240
               Locked          =   -1  'True
               TabIndex        =   49
               Top             =   600
               Width           =   2715
            End
            Begin VB.TextBox txtVersion 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3240
               Locked          =   -1  'True
               TabIndex        =   48
               Top             =   1320
               Width           =   2715
            End
            Begin VB.TextBox txtCompanyName 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3240
               Locked          =   -1  'True
               TabIndex        =   47
               Top             =   1680
               Width           =   3435
            End
            Begin VB.TextBox txtComments 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3240
               Locked          =   -1  'True
               TabIndex        =   46
               Top             =   2040
               Width           =   3435
            End
            Begin VB.TextBox txtMessageFormat 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3240
               Locked          =   -1  'True
               TabIndex        =   45
               Top             =   960
               Width           =   2715
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Details:"
               Height          =   210
               Left            =   1920
               TabIndex        =   57
               Top             =   210
               Width           =   525
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Instrument ID:"
               Height          =   210
               Left            =   2220
               TabIndex        =   56
               Top             =   660
               Width           =   975
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Version:"
               Height          =   210
               Left            =   2580
               TabIndex        =   55
               Top             =   1350
               Width           =   615
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Company Name:"
               Height          =   210
               Left            =   2040
               TabIndex        =   54
               Top             =   1710
               Width           =   1170
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Comments:"
               Height          =   210
               Left            =   2400
               TabIndex        =   53
               Top             =   2070
               Width           =   795
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Message Format:"
               Height          =   210
               Left            =   1980
               TabIndex        =   52
               Top             =   1020
               Width           =   1245
            End
            Begin VB.Shape Shape1 
               BackStyle       =   1  'Opaque
               Height          =   2145
               Left            =   1920
               Top             =   450
               Width           =   5085
            End
         End
         Begin VB.Frame Frame7 
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
            ForeColor       =   &H80000008&
            Height          =   2475
            Left            =   -74970
            TabIndex        =   40
            Top             =   1590
            Width           =   7095
            Begin VB.TextBox txtVSComments 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3330
               Locked          =   -1  'True
               TabIndex        =   61
               Top             =   1410
               Width           =   3435
            End
            Begin VB.TextBox txtVSCompanyName 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3330
               Locked          =   -1  'True
               TabIndex        =   60
               Top             =   1050
               Width           =   3435
            End
            Begin VB.TextBox txtVSVersion 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3330
               Locked          =   -1  'True
               TabIndex        =   59
               Top             =   690
               Width           =   2715
            End
            Begin VB.CommandButton cmdVerSaSettings 
               Caption         =   "VerSa Settings..."
               Enabled         =   0   'False
               Height          =   435
               Left            =   5400
               TabIndex        =   58
               Top             =   1800
               Width           =   1485
            End
            Begin VB.CheckBox chkVersa 
               Appearance      =   0  'Flat
               Caption         =   "Enable Plugin"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   90
               TabIndex        =   42
               Top             =   180
               Width           =   1785
            End
            Begin VB.FileListBox File2 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   1920
               Left            =   120
               Pattern         =   "VS_*.dll"
               TabIndex        =   41
               Top             =   450
               Width           =   1785
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Comments:"
               Height          =   210
               Left            =   2490
               TabIndex        =   64
               Top             =   1440
               Width           =   795
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Company Name:"
               Height          =   210
               Left            =   2130
               TabIndex        =   63
               Top             =   1080
               Width           =   1170
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Version:"
               Height          =   210
               Left            =   2670
               TabIndex        =   62
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Details:"
               Height          =   210
               Left            =   1920
               TabIndex        =   43
               Top             =   210
               Width           =   525
            End
            Begin VB.Shape Shape3 
               BackStyle       =   1  'Opaque
               Height          =   1905
               Left            =   1920
               Top             =   450
               Width           =   5085
            End
         End
         Begin VB.CommandButton cmdSetPlugins 
            Caption         =   "Set Plugins"
            Height          =   615
            Left            =   4020
            TabIndex        =   18
            Top             =   3360
            Width           =   3015
         End
         Begin VB.Frame Frame4 
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
            ForeColor       =   &H80000008&
            Height          =   1545
            Left            =   2640
            TabIndex        =   33
            Top             =   360
            Width           =   4395
            Begin VB.CommandButton cmdHL7Settings 
               Caption         =   "Options..."
               Height          =   255
               Left            =   2310
               TabIndex        =   9
               Top             =   1230
               Width           =   1965
            End
            Begin VB.CheckBox chkHL7 
               Appearance      =   0  'Flat
               Caption         =   "Save HL7 format to disk."
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   210
               TabIndex        =   8
               Top             =   1200
               Width           =   2115
            End
            Begin VB.CheckBox chkUseCTime 
               Appearance      =   0  'Flat
               Caption         =   "Use Current Time on Result Processing."
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   210
               TabIndex        =   7
               Top             =   990
               Width           =   3615
            End
            Begin VB.CheckBox chkAutoConnect 
               Appearance      =   0  'Flat
               Caption         =   "Auto Connect Host"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   210
               TabIndex        =   4
               Top             =   180
               Width           =   2625
            End
            Begin VB.CheckBox chkDateOnly 
               Appearance      =   0  'Flat
               Caption         =   "Do not save time on Result Processing"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   210
               TabIndex        =   6
               Top             =   720
               Width           =   3375
            End
            Begin VB.CheckBox chkConDetect 
               Appearance      =   0  'Flat
               Caption         =   "Detect Instrument Connection"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   210
               TabIndex        =   5
               Top             =   450
               Width           =   2625
            End
         End
         Begin VB.Frame Frame3 
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
            ForeColor       =   &H80000008&
            Height          =   3585
            Left            =   -74910
            TabIndex        =   31
            Top             =   360
            Width           =   7095
            Begin VB.CheckBox chkdisableSave 
               Appearance      =   0  'Flat
               Caption         =   "Disable Save to Database"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   150
               TabIndex        =   118
               Top             =   3120
               Width           =   2325
            End
            Begin VB.CheckBox chkDisableSearch 
               Appearance      =   0  'Flat
               Caption         =   "Disable Search"
               ForeColor       =   &H80000008&
               Height          =   210
               Left            =   150
               TabIndex        =   117
               Top             =   2790
               Width           =   1725
            End
            Begin VB.CommandButton cmdDTPsettings 
               Caption         =   "DTP Settings..."
               Enabled         =   0   'False
               Height          =   405
               Left            =   5070
               TabIndex        =   95
               Top             =   2100
               Width           =   1635
            End
            Begin VB.TextBox txtDTPcomments 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3240
               Locked          =   -1  'True
               TabIndex        =   91
               Top             =   1590
               Width           =   3435
            End
            Begin VB.TextBox txtDTPcompany 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3240
               Locked          =   -1  'True
               TabIndex        =   90
               Top             =   1230
               Width           =   3435
            End
            Begin VB.TextBox txtDTPversion 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   3240
               Locked          =   -1  'True
               TabIndex        =   89
               Top             =   870
               Width           =   2715
            End
            Begin VB.FileListBox File1 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   2130
               Left            =   90
               Pattern         =   "DTP_*.dll"
               TabIndex        =   35
               Top             =   450
               Width           =   1785
            End
            Begin VB.CheckBox chkDatran 
               Appearance      =   0  'Flat
               Caption         =   "Enable Plugin"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   90
               TabIndex        =   34
               Top             =   180
               Width           =   1785
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Comments:"
               Height          =   210
               Left            =   2400
               TabIndex        =   94
               Top             =   1620
               Width           =   795
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Company Name:"
               Height          =   210
               Left            =   2040
               TabIndex        =   93
               Top             =   1260
               Width           =   1170
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Version:"
               Height          =   210
               Left            =   2580
               TabIndex        =   92
               Top             =   900
               Width           =   615
            End
            Begin VB.Shape Shape2 
               BackStyle       =   1  'Opaque
               Height          =   2145
               Left            =   1920
               Top             =   450
               Width           =   5085
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Details:"
               Height          =   210
               Left            =   1920
               TabIndex        =   36
               Top             =   210
               Width           =   525
            End
         End
         Begin VB.TextBox txtProfileID 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1470
            TabIndex        =   2
            Top             =   1140
            Width           =   735
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            Caption         =   "Settings Details"
            ForeColor       =   &H80000008&
            Height          =   2115
            Left            =   90
            TabIndex        =   26
            Top             =   1860
            Width           =   3825
            Begin VB.TextBox txtActivationDate 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1500
               Locked          =   -1  'True
               TabIndex        =   13
               Top             =   1320
               Width           =   2235
            End
            Begin VB.TextBox txtExpirationDate 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1500
               Locked          =   -1  'True
               TabIndex        =   14
               Top             =   1680
               Width           =   2235
            End
            Begin VB.TextBox txtSectionCode 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1500
               Locked          =   -1  'True
               TabIndex        =   12
               Top             =   960
               Width           =   2235
            End
            Begin VB.TextBox txtInstrumentIDdetails 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1500
               Locked          =   -1  'True
               TabIndex        =   10
               Top             =   240
               Width           =   2235
            End
            Begin VB.TextBox txtMessageFormatDetails 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1500
               Locked          =   -1  'True
               TabIndex        =   11
               Top             =   600
               Width           =   2235
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Expiration Date:"
               Height          =   210
               Left            =   300
               TabIndex        =   38
               Top             =   1680
               Width           =   1125
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Activation Date:"
               Height          =   210
               Left            =   330
               TabIndex        =   37
               Top             =   1320
               Width           =   1140
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Section Code:"
               Height          =   210
               Left            =   480
               TabIndex        =   32
               Top             =   1020
               Width           =   1005
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Instrument ID:"
               Height          =   210
               Left            =   480
               TabIndex        =   29
               Top             =   300
               Width           =   975
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Message Format:"
               Height          =   210
               Left            =   240
               TabIndex        =   28
               Top             =   660
               Width           =   1245
            End
         End
         Begin VB.TextBox txtDiffCountTotal 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1470
            TabIndex        =   0
            Top             =   480
            Width           =   735
         End
         Begin VB.Frame Frame6 
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
            ForeColor       =   &H80000008&
            Height          =   1425
            Left            =   3990
            TabIndex        =   39
            Top             =   1860
            Width           =   3045
            Begin VB.TextBox txtSearchCount 
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
               Left            =   1650
               TabIndex        =   120
               Top             =   1020
               Width           =   1215
            End
            Begin VB.CheckBox Check17 
               Appearance      =   0  'Flat
               Caption         =   "Save Non-Numeric Result"
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   210
               TabIndex        =   15
               Top             =   150
               Width           =   2655
            End
            Begin VB.CheckBox chkAutoprint 
               Appearance      =   0  'Flat
               Caption         =   "Auto Print"
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   210
               TabIndex        =   17
               Top             =   690
               Width           =   1485
            End
            Begin VB.CheckBox chkAutoprocess 
               Appearance      =   0  'Flat
               Caption         =   "Auto Process"
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   210
               TabIndex        =   16
               Top             =   420
               Width           =   1425
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Search Result Limit:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   210
               TabIndex        =   121
               Top             =   1050
               Width           =   1425
            End
         End
         Begin VB.Frame Frame10 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   2145
            Left            =   -71280
            TabIndex        =   99
            Top             =   2040
            Width           =   3465
            Begin VB.TextBox txtTemp 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   120
               TabIndex        =   102
               Top             =   1740
               Visible         =   0   'False
               Width           =   1905
            End
            Begin VB.ListBox List1 
               Appearance      =   0  'Flat
               Height          =   1500
               ItemData        =   "frmSettings2.frx":00AC
               Left            =   120
               List            =   "frmSettings2.frx":00AE
               Sorted          =   -1  'True
               TabIndex        =   101
               Top             =   180
               Width           =   3225
            End
            Begin VB.CommandButton cmdSet 
               Caption         =   "Set"
               Height          =   375
               Left            =   2070
               TabIndex        =   100
               Top             =   1710
               Width           =   1275
            End
         End
         Begin RichTextLib.RichTextBox txtMessage 
            Height          =   1275
            Left            =   -74880
            TabIndex        =   105
            Top             =   720
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   2249
            _Version        =   393217
            Appearance      =   0
            TextRTF         =   $"frmSettings2.frx":00B0
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diff Total Format:"
            Height          =   210
            Left            =   210
            TabIndex        =   119
            Top             =   840
            Width           =   1230
         End
         Begin VB.Label Label33 
            Caption         =   "Show Warning at "
            Height          =   255
            Left            =   -74880
            TabIndex        =   108
            Top             =   2130
            Width           =   1335
         End
         Begin VB.Label Label32 
            Caption         =   "samples."
            Height          =   255
            Left            =   -72960
            TabIndex        =   107
            Top             =   2130
            Width           =   735
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Format:"
            Height          =   210
            Left            =   -74730
            TabIndex        =   86
            Top             =   1380
            Width           =   540
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Location:"
            Height          =   210
            Left            =   -74820
            TabIndex        =   85
            Top             =   1020
            Width           =   660
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default Class ID:"
            Height          =   210
            Left            =   240
            TabIndex        =   82
            Top             =   1500
            Width           =   1185
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Default Profile ID:"
            Height          =   210
            Left            =   210
            TabIndex        =   30
            Top             =   1170
            Width           =   1230
         End
         Begin VB.Label txtDetails 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   105
            Left            =   150
            TabIndex        =   27
            Top             =   3060
            Width           =   75
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diff Count Total:"
            Height          =   210
            Left            =   270
            TabIndex        =   25
            Top             =   510
            Width           =   1155
         End
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const REG_SZ As Long = 1 'REG_SZ represents a fixed-length text string.
Private Const REG_BINARY As Long = 3
Private Const REG_DWORD As Long = 4 'REG_DWORD represents data by a number that is 4 bytes long.

Private Const HKEY_CLASSES_ROOT = &H80000000 'The information stored here ensures that the correct program opens when you open a file by using Windows Explorer.
Private Const HKEY_CURRENT_USER = &H80000001 'Contains the root of the configuration information for the user who is currently logged on.
Private Const HKEY_LOCAL_MACHINE = &H80000002 'Contains configuration information particular to the computer (for any user).
Private Const HKEY_USERS = &H80000003 'Contains the root of all user profiles on the computer.

'Return values for all registry functions
Private Const ERROR_SUCCESS = 0
Private Const ERROR_NONE = 0

Private Const KEY_QUERY_VALUE = &H1 'Required to query the values of a registry key.
Private Const KEY_ALL_ACCESS = &H3F 'Combines the STANDARD_RIGHTS_REQUIRED, KEY_QUERY_VALUE, KEY_SET_VALUE, KEY_CREATE_SUB_KEY, KEY_ENUMERATE_SUB_KEYS, KEY_NOTIFY, and KEY_CREATE_LINK access rights.

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long


Private EnablePlugin As Integer
Private EnableVerSa As Integer
Private PlugIN As String
Private VerSa As String
Private DaTran As String
Private Restart As Boolean

Private Enum regkeyFolder
   CLASSES_ROOT = &H80000000
   Current_User = &H80000001
   Local_Machine = &H80000002
   Users = &H80000003
End Enum

Private Enum regkeyDType
   dt_String = 1
   dt_Binary = 3
   dt_DWord = 4
End Enum


Private Sub SaveValue(hKey As regkeyFolder, regType As regkeyDType, strPath As String, strvalue As String, Optional strData As String = "", Optional numData As Long = 0)
Dim ret
RegCreateKey hKey, strPath, ret
If regType = dt_DWord Then
    RegSetValueEx ret, strvalue, 0, regType, numData, Len(numData)
ElseIf regType = dt_Binary Then
    RegSetValueEx ret, strvalue, 0, regType, strData, Len(strData)
ElseIf regType = dt_String Then
    RegSetValueEx ret, strvalue, 0, regType, strData, Len(strData)
End If
RegCloseKey ret
End Sub
 
Private Function QueryValue(myKey1 As regkeyFolder, sKeyName As String, sValueName As String) As String
       
  Dim lRetVal As Long         'result of the API functions
  Dim hKey As Long         'handle of opened key
  Dim vValue As Variant      'setting of queried value

  lRetVal = RegOpenKeyEx(myKey1, sKeyName, 0, KEY_QUERY_VALUE, hKey) 'Open Key to Query a value
  lRetVal = QueryValueEx(hKey, sValueName, vValue) 'Query (determine) the value stored

  QueryValue = vValue 'Set the Form's Caption to whatever text was stored
  RegCloseKey (hKey) 'Close the Key
       
End Function
   
Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
       
       Dim Data As Long
       Dim retVal As Long 'Return value of RegQuery functions
       Dim lType As Long 'Determine data type of present data
       Dim lValue As Long 'Long value
       Dim sValue As String 'String value

       On Error GoTo QueryValueExError

       ' Determine the size and type of data to be read
       retVal = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, Data)
       
       If retVal <> ERROR_NONE Then Error 5

       Select Case lType
           ' Determine strings
           Case REG_SZ:
               sValue = String(Data, 0)

               retVal = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, Data)
               
               If retVal = ERROR_NONE Then
                   vValue = Left$(sValue, Data - 1)
               Else
                   vValue = Empty
               End If
               
           ' Determine DWORDS
           Case REG_DWORD:
               retVal = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, Data)
               
               If retVal = ERROR_NONE Then vValue = lValue
           
           Case Else
               'all other data types not supported
               retVal = -1
       End Select
    
QueryValueExError:
       QueryValueEx = retVal
       Exit Function

   End Function

Private Sub cboScheme_Click()
Select Case cboScheme.Text
    Case "Default Colors"
        picTitleBar.BackColor = &HA66B2B
        picTitleText.BackColor = &HFFFFFF
        picBodyFrame.BackColor = &H6B4B01
        picBodyText.BackColor = &HFFFFFF
    Case "System Colors"
        picTitleBar.BackColor = vbActiveTitleBar
        picTitleText.BackColor = vbActiveTitleBarText
        picBodyFrame.BackColor = vbButtonFace
        picBodyText.BackColor = vbButtonText
    Case "Red/White/Orange/Blue"
        picTitleBar.BackColor = 1376961
        picTitleText.BackColor = 16777215
        picBodyFrame.BackColor = 8047855
        picBodyText.BackColor = 8388736
    Case "Green/Yellow/Red"
        picTitleBar.BackColor = 3512166
        picTitleText.BackColor = 16777215
        picBodyFrame.BackColor = 5636074
        picBodyText.BackColor = 16512
    Case "Black/White/Gray"
        picTitleBar.BackColor = 0
        picTitleText.BackColor = 16777215
        picBodyFrame.BackColor = 6250335
        picBodyText.BackColor = 16777215
    Case "Green/White/Gray"
        picTitleBar.BackColor = 8421376
        picTitleText.BackColor = 16777215
        picBodyFrame.BackColor = 14084030
        picBodyText.BackColor = 6250335
    Case "Red/Black/Yellow/White"
        picTitleBar.BackColor = 255
        picTitleText.BackColor = 65535
        picBodyFrame.BackColor = 986895
        picBodyText.BackColor = 16777215
    Case "Gray/Orange/Red/White"
        picTitleBar.BackColor = 5131854
        picTitleText.BackColor = 8765415
        picBodyFrame.BackColor = 147
        picBodyText.BackColor = 16777215
    Case "Blue/White/Silver/Gray"
        picTitleBar.BackColor = 13866518
        picTitleText.BackColor = 16777215
        picBodyFrame.BackColor = 14737632
        picBodyText.BackColor = 6250335
End Select
End Sub

Private Sub Check1_Click()
If chkRegistry.Value = 1 Then
    SaveValue Current_User, dt_Binary, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", "16777216"
Else
'    RegDeleteKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" & App.Title 'Delete the Key created with API
    SaveValue Current_User, dt_Binary, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", 0
End If
End Sub

Private Sub chkAutoprint_Click()
If chkAutoprint.Value = 1 Then chkAutoprocess.Value = 1
End Sub

Private Sub chkAutoprocess_Click()
    If chkAutoprocess.Value = 0 Then chkAutoprint.Value = 0
End Sub

Private Sub chkDateOnly_Click()
If chkDateOnly.Value = 1 Then chkUseCTime.Value = 0
End Sub

Private Sub chkDatran_Click()
If chkDatran.Value = 1 Then
    File1.Enabled = True
    If Len(txtDTPversion.Text) > 0 Then
        cmdDTPsettings.Enabled = True
    Else
        cmdDTPsettings.Enabled = False
    End If
    chkDisableSearch.Enabled = True
    chkdisableSave.Enabled = True
Else
    File1.Enabled = False
    cmdDTPsettings.Enabled = False
    txtDTPversion.Text = ""
    txtDTPcompany.Text = ""
    txtDTPcomments.Text = ""
    chkDisableSearch.Enabled = False
    chkdisableSave.Enabled = False
End If
End Sub

Private Sub chkExport_Click()
If chkExport.Value = 1 Then
    txtLocation.Enabled = True
    cboExport.Enabled = True
Else
    txtLocation.Enabled = False
    cboExport.Enabled = False
End If
End Sub

Private Sub chkHL7_Click()
If chkHL7.Value = 1 Then cmdHL7Settings.Enabled = True Else cmdHL7Settings.Enabled = False
End Sub

Private Sub chkPlugin_Click()
    If chkPlugin.Value = 1 Then
        fileList.Enabled = True
        setDetails
    Else
        fileList.Enabled = False
        clearDetails
    End If
End Sub

Private Sub chkRegistry_Click()
If chkRegistry.Value = 1 Then
    SaveValue Current_User, dt_DWord, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", , 1
Else
    SaveValue Current_User, dt_DWord, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", , 0
End If
End Sub

Private Sub chkTaskM_Click()
If chkTaskM.Value = 1 Then
    SaveValue Current_User, dt_DWord, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", , 1
Else
    SaveValue Current_User, dt_DWord, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", , 0
End If
End Sub

Private Sub chkUsbStor_Click()
If chkUsbStor.Value = 1 Then
    SaveValue Local_Machine, dt_DWord, "System\CurrentControlSet\Services\USBSTOR", "Start", , 4
Else
    SaveValue Local_Machine, dt_DWord, "System\CurrentControlSet\Services\USBSTOR", "Start", , 3
End If
End Sub

Private Sub chkUseCTime_Click()
If chkUseCTime.Value = 1 Then
    chkDateOnly.Value = 0
End If
End Sub

Private Sub chkVersa_Click()
    If chkVersa.Value = 1 Then
        File2.Enabled = True
        If Len(txtVSVersion.Text) > 0 Then
            cmdVerSaSettings.Enabled = True
        Else
            cmdVerSaSettings.Enabled = False
        End If
    Else
        File2.Enabled = False
        txtVSVersion.Text = ""
        txtVSCompanyName.Text = ""
        txtVSComments.Text = ""
    End If
End Sub

Private Sub cmdBrowse_Click()
    Dim X As Integer
    '-- Cheap way to use the common dialog box as a directory-picker
    X = 3

    cdgExportLocation.CancelError = True        'do not terminate on error

    On Error Resume Next         'I will hande errors

    cdgExportLocation.Action = 1              'Present "open" dialog

    '-- If FileTitle is null, user did not override the default (*.*)
    If cdgExportLocation.FileTitle <> "" Then X = Len(cdgExportLocation.FileTitle)

    If Err = 0 Then
        ChDrive cdgExportLocation.filename
        txtLocation.Text = Left(cdgExportLocation.filename, Len(cdgExportLocation.filename) - X)
    Else
      '-- User pressed "Cancel"
    End If
End Sub

Private Sub initCdgExportLocation()
    With cdgExportLocation
        .Flags = cdlOFNPathMustExist
        .Flags = .Flags Or cdlOFNHideReadOnly
        .Flags = .Flags Or cdlOFNNoChangeDir
        .Flags = .Flags Or cdlOFNExplorer
        .Flags = .Flags Or cdlOFNNoValidate
        .filename = "*.*"
    End With
End Sub

Private Sub cmdClose_Click()
    If Restart Then
        If isResultVisible Then frmSearchResult.Visible = False
        InitPlugin
        If isResultVisible Then frmSearchResult.Visible = False
    End If
    InitPlugin
    Unload Me
End Sub

Private Sub cmdDTPsettings_Click()
Dim mDef As Object
Set mDef = CreateObject(Replace(Format(File1.filename, ">"), ".DLL", ".MessageDef"))
mDef.showsettings
End Sub

Private Sub cmdHL7Settings_Click()
    frmHL7Settings.Show 1
End Sub

Private Sub cmdPluginSettings_Click()
Dim mDef As Object
On Error GoTo errNum
    Set mDef = CreateObject(Replace(Format(fileList.filename, ">"), ".DLL", ".MessageDef"))
    mDef.showsettings
errNum:
If Err.Number = 438 Then
    MsgBox "Not Supported", vbOKOnly + vbInformation
End If
End Sub

Private Sub cmdRegistry_Click()
SaveValue Current_User, dt_DWord, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", , 0
chkRegistry.Value = 0
Shell "regedit.exe"
End Sub

Private Sub cmdSave_Click()
saveSettings
If frmSearch.Visible Then
    setFormColors frmSearch
    frmSearch.chkResultDate.BackColor = picBodyFrame.BackColor
    frmSearch.chkResultDate.ForeColor = picBodyText.BackColor
End If
setFormColors frmControlPanel
If isConnected Then setFormColors frmReceiveData
If isResultVisible Then setFormColors frmSearchResult
setFormColors Me
sstSettings.BackColor = fraBody(0).BackColor
If sstSettings.TabVisible(1) = True Then
    Restart = True
End If
sstSettings.TabVisible(1) = False
sstSettings.TabVisible(2) = False
sstSettings.TabVisible(3) = False
sstSettings.TabVisible(5) = False
EnablePlugin = Val(GetSetting("Lab-I", "Settings", "EnablePlugin", "0"))
If EnablePlugin = 1 Then
    txtInstrumentIDdetails.Text = txtInstrumentID
    txtMessageFormatDetails.Text = txtMessageFormat
Else
    txtInstrumentIDdetails.Text = "HMX"
    txtMessageFormatDetails.Text = "STKS"
    txtSectionCode.Text = "HEMA"
End If
MsgBox "Settings has been successfully saved.", vbOKOnly, App.ProductName
End Sub

Private Sub cmdSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 2 Then
        SaveRegToFile HKEY_CURRENT_USER, "SOFTWARE\VB and VBA Program Settings", IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\InterfaceRegistry.reg"
    End If
End Sub

Private Sub cmdSet_Click()
Dim sourceFile As String
Dim arrColor() As String
On Error Resume Next
    If Len(Trim(txtTemp)) > 0 Then
        arrColor = Split(txtTemp.Text, "|")
        picBodyFrame.BackColor = Val(arrColor(0))
        picBodyText.BackColor = Val(arrColor(1))
        picTitleBar.BackColor = Val(arrColor(2))
        picTitleText.BackColor = Val(arrColor(3))
    End If
    sourceFile = App.Path & "\Themes\" & Trim(List1.Text) & "\lab-i_L.sys"
    fsoFiles.DeleteFile App.Path & "\lab-i_L.sys", True
    fsoFiles.CopyFile sourceFile, App.Path & "\lab-i_L.sys", True
    If Screen.Width = 12000 Then
        If (Val(Format(Date, "mm")) = 12 And Val(Format(Date, "dd")) >= 25 And Val(Format(Date, "dd")) <= 31) Or Format(Date, "mmdd") = "0101" Then
            mdiMain.Picture = LoadPicture(App.Path & "\" & "lab-ix.sys")
        Else
            mdiMain.Picture = LoadPicture(App.Path & "\" & "lab-i.sys")
        End If
    Else
        If (Val(Format(Date, "mm")) = 12 And Val(Format(Date, "dd")) >= 25 And Val(Format(Date, "dd")) <= 31) Or Format(Date, "mmdd") = "0101" Then
            mdiMain.Picture = LoadPicture(App.Path & "\" & "lab-ixl.sys")
        Else
            mdiMain.Picture = LoadPicture(App.Path & "\" & "lab-i_L.sys")
        End If
    End If
End Sub

Private Sub cmdSetPlugins_Click()
    txtDetails_DblClick
End Sub

Private Sub cmdSetting_Click()
    frmOBISetiings.Show 1
End Sub

Private Sub cmdTaskMAn_Click()
SaveValue Current_User, dt_DWord, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", , 0
cmdTaskMAn.Value = 0
Shell "taskmgr.exe"
End Sub

Private Sub cmdTestPlugin_Click()
Dim cForms As New ClassForms
Dim rstTemp As ADODB.Recordset
Set rstTemp = myConnection.Execute("select * from labunprocessed")
    If rstTemp.RecordCount > 0 Then
        frmPlugInTest.Show 1
    Else
        MsgBox "No Record found.", vbOKOnly, App.ProductName
    End If
End Sub

Private Sub cmdVerSaSettings_Click()
Dim mDef As Object
Set mDef = CreateObject(Replace(Format(File2.filename, ">"), ".DLL", ".mFunctions"))
mDef.Settings
End Sub

Private Sub File1_Click()
setDTP_Details True
End Sub

Private Sub File2_Click()
setVerSaDetails True
End Sub

Private Sub fileList_Click()
setDetails True
End Sub

Private Sub setThemeDir()
On Error Resume Next
ListSubDirs App.Path & "\Themes\"
End Sub

Private Sub Form_Load()
setFormColors Me
fillcboScheme
fillcboExport
initCdgExportLocation
sstSettings.BackColor = fraBody(0).BackColor
setThemeDir
Restart = False
sstSettings.TabVisible(1) = False
sstSettings.TabVisible(2) = False
sstSettings.TabVisible(3) = False
sstSettings.TabVisible(5) = False
If IsDate(ActivationDate) Then txtActivationDate.Text = Format(ActivationDate, "dddd, mmmm dd, yyyy") Else txtActivationDate.Text = "Date Not Set"
If IsDate(ExpirationDate) Then txtExpirationDate.Text = Format(ExpirationDate, "dddd, mmmm dd, yyyy") Else txtExpirationDate.Text = "Date Not Set"
fileList.Path = App.Path & "\Plugins"
File1.Path = App.Path & "\Plugins"
File2.Path = App.Path & "\Plugins"
txtDiffCountTotal.Text = GetSetting("Lab-I", "Settings", "DiffCount", "1")
txtDiffFormat.Text = GetSetting("Lab-I", "Settings", "DiffCountFormat", "0")
txtProfileID.Text = GetSetting("Lab-I", "Settings", "ProfileID", "1")
txtClass.Text = GetSetting("Lab-I", "Settings", "ClassID", "")
EnablePlugin = Val(GetSetting("Lab-I", "Settings", "EnablePlugin", "0"))
EnableVerSa = Val(GetSetting("Lab-I", "VerSa", "EnablePlugin", "0"))
chkDatran.Value = Val(GetSetting("Lab-I", "Datran", "EnablePlugin", "0"))
PlugIN = Format(GetSetting("Lab-I", "Settings", "PlugIN", ""), ">")
VerSa = Format(GetSetting("Lab-I", "VerSa", "PlugIN", ""), ">")
DaTran = Format(GetSetting("Lab-I", "DaTran", "PlugIN", ""), ">")
chkPlugin.Value = EnablePlugin
chkVersa.Value = EnableVerSa
chkConDetect.Value = Val(GetSetting("Lab-I", "Settings", "DetectConnection", "1"))
chkDateOnly.Value = Val(GetSetting("Lab-I", "Settings", "DateOnly", "0"))
chkUseCTime.Value = Val(GetSetting("Lab-I", "Settings", "UseCurrentTime", "0"))
chkAutoprocess.Value = Val(GetSetting("Lab-I", "Settings", "Autoprocess", "0"))
chkAutoprint.Value = Val(GetSetting("Lab-I", "Settings", "Autoprint", "0"))
chkAutoConnect.Value = Val(GetSetting("Lab-I", "Settings", "AutoConnect", "0"))
Check17.Value = Val(GetSetting("HostForms_SA", "Settings", "SaveStringResult", "0"))
chkExport.Value = Val(GetSetting("Lab-I", "VerSa", "EnableExport", "0"))
txtLocation.Text = GetSetting("Lab-I", "VerSa", "ExportLocation", "")
findcboType GetSetting("Lab-I", "VerSa", "ExportType", "")
If chkExport.Value = 1 Then
    txtLocation.Enabled = True
    cboExport.Enabled = True
Else
    txtLocation.Enabled = False
    cboExport.Enabled = False
End If
If EnablePlugin = 1 Then
    fileList.Enabled = True
    setDetails
    txtInstrumentIDdetails.Text = txtInstrumentID
    txtMessageFormatDetails.Text = txtMessageFormat
Else
    fileList.Enabled = False
    clearDetails
    txtInstrumentIDdetails.Text = "HMX"
    txtMessageFormatDetails.Text = "STKS"
    txtSectionCode.Text = "HEMA"
End If
If EnableVerSa Then
    setVerSaDetails
    File2.Enabled = True
Else
    File2.Enabled = False
End If

If chkDatran.Value = 1 Then
    setDTP_Details
    File1.Enabled = True
    chkDisableSearch.Enabled = True
    chkdisableSave.Enabled = True
Else
    File1.Enabled = False
    chkDisableSearch.Enabled = False
    chkdisableSave.Enabled = False
End If

If GetSetting("Lab-I", "Settings", "Method", "B") = "A" Then optMethodA.Value = True Else optMethodB.Value = True

picTitleBar.BackColor = GetSetting("Lab-I", "FormColors", "TitleBar", &HA66B2B)
picTitleText.BackColor = GetSetting("Lab-I", "FormColors", "TitleText", &HFFFFFF)
picBodyFrame.BackColor = GetSetting("Lab-I", "FormColors", "BodyFrame", &H80000008)
picBodyText.BackColor = GetSetting("Lab-I", "FormColors", "BodyText", &HFFFFFF)
picSaved.BackColor = GetSetting("Lab-I", "FormColors", "SavedDocCol", picSaved.BackColor)
picPrinted.BackColor = GetSetting("Lab-I", "FormColors", "PrintedDocCol", picPrinted.BackColor)
chkWarning.Value = Val(GetSetting("Lab-I", "Others", "ENABLEWARNING", "0"))
txtMessage.Text = GetSetting("Lab-I", "Others", "WARNINGMESSAGE", "")
txtSamples.Text = GetSetting("Lab-I", "Others", "WARNINGCOUNT", "0")
txtSearchCount.Text = GetSetting("Lab-I", "Settings", "SearchCount", "300")
chkDisableSearch.Value = Val(GetSetting("Lab-I", "DaTran", "DisableSearch", "0"))
chkdisableSave.Value = Val(GetSetting("Lab-I", "DaTran", "DisableSave", "0"))

'Registry
chkRegistry.Value = Val(QueryValue(Current_User, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"))
chkTaskM.Value = Val(QueryValue(Current_User, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr"))
Check1.Value = Val(QueryValue(Current_User, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions"))
If Val(QueryValue(Local_Machine, "System\CurrentControlSet\Services\USBSTOR", "Start")) = 3 Then
    chkUsbStor.Value = 0
Else
    chkUsbStor.Value = 1
End If
chkHL7.Value = Val(GetSetting("Lab-I", "HL7", "Enable", "0"))
If chkHL7.Value = 1 Then cmdHL7Settings.Enabled = True Else cmdHL7Settings.Enabled = False
End Sub

Private Sub findcboType(myvalue As String)
Dim ctr As Integer
Dim mFound As Boolean
Dim tIndex As Integer
mFound = False
With cboExport
    For ctr = 0 To .ListCount
        .ListIndex = ctr
        If .Text = myvalue Then
            mFound = True
            tIndex = .ListIndex
            Exit For
        End If
    Next ctr
If mFound Then .ListIndex = tIndex
End With
End Sub

Private Sub setDetails(Optional setSelected As Boolean = False)
Dim mDef As Object
Dim ctr As Integer
If Not setSelected Then
    For ctr = 0 To fileList.ListCount - 1
        If Format(fileList.List(ctr), ">") = PlugIN Then
            Set mDef = CreateObject(Replace(PlugIN, ".DLL", ".MessageDef"))
            fileList.Selected(ctr) = True
            txtInstrumentID.Text = mDef.InstrumentID
            txtVersion.Text = mDef.Version
            txtCompanyName.Text = mDef.CompanyName
            txtComments.Text = mDef.Comments
            txtMessageFormat.Text = mDef.MessageFormat
            txtSectionCode.Text = mDef.SectionCode
            Exit For
        End If
    Next ctr
Else
    Set mDef = CreateObject(Replace(Format(fileList.filename, ">"), ".DLL", ".MessageDef"))
    txtInstrumentID.Text = mDef.InstrumentID
    txtVersion.Text = mDef.Version
    txtCompanyName.Text = mDef.CompanyName
    txtComments.Text = mDef.Comments
    txtMessageFormat.Text = mDef.MessageFormat
    txtSectionCode.Text = mDef.SectionCode
End If
End Sub

Private Sub setDTP_Details(Optional setSelected As Boolean = False)
Dim mDef As Object
Dim ctr As Integer
If Not setSelected Then
    For ctr = 0 To File1.ListCount - 1
        If Format(File1.List(ctr), ">") = DaTran Then
            Set mDef = CreateObject(Replace(DaTran, ".DLL", ".MessageDef"))
            File1.Selected(ctr) = True
            txtDTPversion.Text = mDef.Version
            txtDTPcompany.Text = mDef.CompanyName
            txtDTPcomments.Text = mDef.Comments
            Exit For
        End If
    Next ctr
Else
    Set mDef = CreateObject(Replace(Format(File1.filename, ">"), ".DLL", ".MessageDef"))
    txtDTPversion.Text = mDef.Version
    txtDTPcompany.Text = mDef.CompanyName
    txtDTPcomments.Text = mDef.Comments
End If
End Sub

Private Sub setVerSaDetails(Optional setSelected As Boolean = False)
Dim mDef As Object
Dim ctr As Integer
If Not setSelected Then
    For ctr = 0 To File2.ListCount - 1
        If Format(File2.List(ctr), ">") = VerSa Then
            Set mDef = CreateObject(Replace(VerSa, ".DLL", ".mFunctions"))
            File2.Selected(ctr) = True
            txtVSVersion.Text = mDef.Version
            txtVSCompanyName.Text = mDef.CompanyName
            txtComments.Text = mDef.Comments
            Exit For
        End If
    Next ctr
Else
    Set mDef = CreateObject(Replace(Format(File2.filename, ">"), ".DLL", ".mFunctions"))
    txtVSVersion.Text = mDef.Version
    txtVSCompanyName.Text = mDef.CompanyName
    txtVSComments.Text = mDef.Comments
End If
End Sub

Private Sub clearDetails()
    txtInstrumentID.Text = ""
    txtVersion.Text = ""
    txtCompanyName.Text = ""
    txtComments.Text = ""
    txtMessageFormat.Text = ""
End Sub

Private Sub saveSettings()
Dim mDirectory As String
If optMethodA.Value = True Then SaveSetting "Lab-I", "Settings", "Method", "A" Else SaveSetting "Lab-I", "Settings", "Method", "B"
SaveSetting "Lab-I", "Settings", "DiffCount", txtDiffCountTotal.Text
SaveSetting "Lab-I", "Settings", "DiffCountFormat", txtDiffFormat.Text
SaveSetting "Lab-I", "Settings", "ClassID", txtClass.Text
SaveSetting "Lab-I", "Settings", "SearchCount", txtSearchCount.Text
SaveSetting "Lab-I", "Settings", "ProfileID", txtProfileID.Text
SaveSetting "Lab-I", "Settings", "EnablePlugin", Trim(Str(chkPlugin.Value))
SaveSetting "Lab-I", "Settings", "DetectConnection", Str(chkConDetect.Value)
SaveSetting "Lab-I", "Settings", "DateOnly", Str(chkDateOnly.Value)
SaveSetting "Lab-I", "Settings", "UseCurrentTime", Str(chkUseCTime.Value)
SaveSetting "Lab-I", "Settings", "Autoprocess", Trim(Str(chkAutoprocess.Value))
SaveSetting "Lab-I", "Settings", "Autoprint", Trim(Str(chkAutoprint.Value))
SaveSetting "HostForms_SA", "Settings", "SaveStringResult", Trim(Str(Check17.Value))
SaveSetting "Lab-I", "VerSa", "EnablePlugin", Trim(Str(chkVersa.Value))
SaveSetting "Lab-I", "DaTran", "EnablePlugin", Trim(Str(chkDatran.Value))
SaveSetting "Lab-I", "Settings", "AutoConnect", Trim(Str(chkAutoConnect))
If chkPlugin.Value = 1 And Len(Trim(fileList.filename)) > 0 Then
    SaveSetting "Lab-I", "Settings", "PlugIN", fileList.filename
Else
    SaveSetting "Lab-I", "Settings", "PlugIN", ""
End If
If chkVersa.Value = 1 And Len(Trim(File2.filename)) > 0 Then
    SaveSetting "Lab-I", "VerSa", "PlugIN", File2.filename
Else
    SaveSetting "Lab-I", "VerSa", "PlugIN", ""
End If
If chkDatran.Value = 1 And Len(Trim(File1.filename)) > 0 Then
    SaveSetting "Lab-I", "DaTran", "PlugIN", File1.filename
Else
    SaveSetting "Lab-I", "DaTran", "PlugIN", ""
End If
SaveSetting "Lab-I", "FormColors", "TitleBar", picTitleBar.BackColor
SaveSetting "Lab-I", "FormColors", "TitleText", picTitleText.BackColor
SaveSetting "Lab-I", "FormColors", "BodyFrame", picBodyFrame.BackColor
SaveSetting "Lab-I", "FormColors", "BodyText", picBodyText.BackColor
SaveSetting "Lab-I", "FormColors", "SavedDocCol", picSaved.BackColor
SaveSetting "Lab-I", "FormColors", "PrintedDocCol", picPrinted.BackColor
SaveSetting "Lab-I", "VerSa", "EnableExport", Trim(Str(chkExport.Value))
SaveSetting "Lab-I", "VerSa", "ExportLocation", Trim(txtLocation.Text)
SaveSetting "Lab-I", "VerSa", "ExportType", Trim(cboExport.Text)
SaveSetting "Lab-I", "Others", "ENABLEWARNING", chkWarning.Value
SaveSetting "Lab-I", "Others", "WARNINGMESSAGE", txtMessage.Text
SaveSetting "Lab-I", "Others", "WARNINGCOUNT", txtSamples.Text
SaveSetting "Lab-I", "DaTran", "DisableSearch", chkDisableSearch.Value
SaveSetting "Lab-I", "DaTran", "DisableSave", chkdisableSave.Value
SaveSetting "Lab-I", "HL7", "Enable", chkHL7.Value

mDirectory = Replace(Format(DBDatabase, ">"), "LABORATORY.MDB", "")
If Len(Dir(mDirectory & "\" & Environ("USERNAME") & "@" & Environ("USERDOMAIN"), vbDirectory)) = 0 Then
    fsoFiles.Createfolder mDirectory & "\" & Environ("USERNAME") & "@" & Environ("USERDOMAIN")
End If
mDirectory = mDirectory & "\" & Environ("USERNAME") & "@" & Environ("USERDOMAIN")
SaveRegToFile HKEY_CURRENT_USER, "SOFTWARE\VB and VBA Program Settings", mDirectory & "\Laboratory.reg"
SaveSetting "Lab-I", "Settings", "LastUser", Environ("USERNAME") & "@" & Environ("USERDOMAIN")
End Sub

Private Sub List1_Click()
On Error Resume Next
If Not LoadFileToTB(txtTemp, App.Path & "\Themes\" & Trim(List1.Text) & "\Theme.lit") Then
    txtTemp.Text = ""
End If
End Sub

Private Sub picBodyFrame_Click()
    cdgColor.Color = picTitleBar.BackColor
    cdgColor.ShowColor
    picBodyFrame.BackColor = cdgColor.Color
End Sub

Private Sub picBodyText_Click()
    cdgColor.Color = picTitleBar.BackColor
    cdgColor.ShowColor
    picBodyText.BackColor = cdgColor.Color
End Sub

Private Sub picPrinted_Click()
    cdgColor.Color = picPrinted.BackColor
    cdgColor.ShowColor
    picPrinted.BackColor = cdgColor.Color
End Sub

Private Sub picSaved_Click()
    cdgColor.Color = picSaved.BackColor
    cdgColor.ShowColor
    picSaved.BackColor = cdgColor.Color
End Sub

Private Sub picTitleBar_Click()
    cdgColor.Color = picTitleBar.BackColor
    cdgColor.ShowColor
    picTitleBar.BackColor = cdgColor.Color
End Sub

Private Sub picTitleText_Click()
    cdgColor.Color = picTitleBar.BackColor
    cdgColor.ShowColor
    picTitleText.BackColor = cdgColor.Color
End Sub

Private Sub sstSettings_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
    
    End If
End Sub

Private Sub txtDetails_DblClick()
    Dim cForms As New ClassForms
    If sstSettings.TabVisible(1) = True Then
        sstSettings.TabVisible(1) = False
        sstSettings.TabVisible(2) = False
        sstSettings.TabVisible(3) = False
        sstSettings.TabVisible(5) = False
    Else
        If cForms.showInputBox("Password", , , , , , True) = "standard" Then
            sstSettings.TabVisible(1) = True
            sstSettings.TabVisible(2) = True
            sstSettings.TabVisible(3) = True
            sstSettings.TabVisible(5) = True
        End If
    End If
End Sub

Private Sub fillcboScheme()
With cboScheme
    .AddItem "Default Colors"
    .AddItem "System Colors"
    .AddItem "Red/White/Orange/Blue"
    .AddItem "Green/Yellow/Red"
    .AddItem "Black/White/Gray"
    .AddItem "Green/White/Gray"
    .AddItem "Red/Black/Yellow/White"
    .AddItem "Gray/Orange/Red/White"
    .AddItem "Blue/White/Silver/Gray"
End With
End Sub

Private Sub fillcboExport()
With cboExport
    .AddItem ""
    .AddItem "Comma Separated Values (CSV)"
    .AddItem "Excel (XLS)"
    .AddItem "Rich Text Format (RTF)"
    .AddItem "HTML"
    .AddItem "Word for Windows (DOC)"
    .AddItem "Text File"
    .AddItem "PDF File (PDF)"
    .ListIndex = 0
End With
End Sub

Private Sub txtDTPversion_Change()
If Len(txtDTPversion.Text) > 0 Then
    cmdDTPsettings.Enabled = True
Else
    cmdDTPsettings.Enabled = False
End If
End Sub

Private Sub txtVSVersion_Change()
If Len(txtVSVersion.Text) > 0 Then
    cmdVerSaSettings.Enabled = True
Else
    cmdVerSaSettings.Enabled = False
End If
End Sub

Private Sub ListSubDirs(Path)
    On Error Resume Next
    Dim Count, D(), i, DirName      ' Declare variables.
    DirName = Dir(Path, 16)         ' Get first directory name.
    Do While DirName <> ""
        ' A file or directory name was returned
        If DirName <> "." And DirName <> ".." Then
            ' Not a parent or current directory entry so process it
            If GetAttr(Path + DirName) = 16 Then
                ' This is a directory
                If (Count Mod 10) = 0 Then
                    ' Resize the array
                    ReDim Preserve D(Count + 10)
                End If
                Count = Count + 1   ' Increment counter.
                D(Count) = DirName  ' Add directory name to array
            End If
        End If
        DirName = Dir   ' Get another directory name.
    Loop
    For i = 1 To Count
        ' Put each directory's name in the list box.
        List1.AddItem D(i)
        ' The following is the recursive call
        ' Find this directory's sub-directories
        ListSubDirs Path & D(i) & "\"
    Next i
    DoEvents
End Sub

Public Function LoadFileToTB(TxtBox As Object, FilePath As _
   String, Optional Append As Boolean = False) As Boolean
   
'PURPOSE: Loads file specified by FilePath into textcontrol
'(e.g., Text Box, Rich Text Box) specified by TxtBox

'If Append = true, then loaded text is appended to existing
' contents else existing contents are overwritten

'Returns: True if Successful, false otherwise

Dim iFile As Integer
Dim s As String

If Dir(FilePath) = "" Then Exit Function

On Error GoTo errorhandler:
s = TxtBox.Text

iFile = FreeFile
Open FilePath For Input As #iFile
s = Input(LOF(iFile), #iFile)
If Append Then
    TxtBox.Text = TxtBox.Text & s
Else
    TxtBox.Text = s
End If

LoadFileToTB = True

errorhandler:
If iFile > 0 Then Close #iFile

End Function

