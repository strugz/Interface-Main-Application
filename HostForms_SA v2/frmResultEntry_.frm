VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmResultEntry 
   BackColor       =   &H00400040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Result Entry and Verification"
   ClientHeight    =   9045
   ClientLeft      =   4785
   ClientTop       =   2100
   ClientWidth     =   11370
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00A66B2B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   60
      TabIndex        =   33
      Top             =   -60
      Width           =   11295
      Begin VB.CommandButton cmdDelta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6720
         TabIndex        =   80
         Top             =   150
         Visible         =   0   'False
         Width           =   465
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
         Left            =   8610
         Locked          =   -1  'True
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   150
         Width           =   2565
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
         Left            =   7320
         TabIndex        =   76
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   73
         Top             =   180
         Width           =   1680
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result Entry and Verification"
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
         Left            =   3630
         TabIndex        =   35
         Top             =   180
         Visible         =   0   'False
         Width           =   2790
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2145
      Index           =   0
      Left            =   60
      TabIndex        =   34
      Top             =   510
      Width           =   11295
      Begin VB.Frame fraBody 
         BackColor       =   &H006B4B01&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1965
         Index           =   6
         Left            =   30
         TabIndex        =   61
         Top             =   150
         Width           =   3615
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
            TabIndex        =   6
            Top             =   660
            Width           =   2565
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
            TabIndex        =   20
            Top             =   1650
            Width           =   375
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
            Left            =   2430
            MaxLength       =   6
            TabIndex        =   10
            Top             =   990
            Width           =   1215
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
            TabIndex        =   9
            Top             =   990
            Width           =   975
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
            MaxLength       =   50
            TabIndex        =   15
            Top             =   1320
            Width           =   2295
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
            MaxLength       =   50
            TabIndex        =   3
            Top             =   330
            Width           =   2565
         End
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
            TabIndex        =   0
            Top             =   0
            Width           =   2565
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
            TabIndex        =   71
            Top             =   1320
            Width           =   285
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
            TabIndex        =   70
            Top             =   1650
            Width           =   285
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
            TabIndex        =   69
            Top             =   690
            Width           =   990
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
            TabIndex        =   68
            Top             =   30
            Width           =   795
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
            TabIndex        =   67
            Top             =   1650
            Width           =   1905
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
            TabIndex        =   66
            Top             =   1680
            Width           =   960
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
            TabIndex        =   65
            Top             =   1020
            Width           =   330
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
            TabIndex        =   64
            Top             =   1020
            Width           =   975
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
            TabIndex        =   63
            Top             =   1350
            Width           =   975
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
            TabIndex        =   62
            Top             =   360
            Width           =   990
         End
      End
      Begin VB.Frame fraBody 
         BackColor       =   &H006B4B01&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1965
         Index           =   4
         Left            =   3720
         TabIndex        =   50
         Top             =   150
         Width           =   3705
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
            TabIndex        =   7
            Top             =   660
            Width           =   2565
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
            TabIndex        =   21
            Top             =   1650
            Width           =   465
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
            TabIndex        =   51
            Top             =   1320
            Visible         =   0   'False
            Width           =   915
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
            TabIndex        =   12
            Top             =   990
            Width           =   885
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
            TabIndex        =   11
            Top             =   990
            Width           =   1005
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
            TabIndex        =   1
            Top             =   0
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
            TabIndex        =   17
            Top             =   1320
            Width           =   2295
         End
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
            MaxLength       =   50
            TabIndex        =   4
            Top             =   330
            Width           =   2565
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
            TabIndex        =   60
            Top             =   1320
            Width           =   285
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
            TabIndex        =   59
            Top             =   1650
            Width           =   285
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
            TabIndex        =   58
            Top             =   660
            Width           =   1020
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
            TabIndex        =   23
            Top             =   1650
            Width           =   1815
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
            TabIndex        =   57
            Top             =   1680
            Width           =   1065
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
            TabIndex        =   56
            Top             =   60
            Width           =   990
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
            TabIndex        =   55
            Top             =   1380
            Width           =   1050
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
            TabIndex        =   54
            Top             =   1020
            Width           =   450
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
            TabIndex        =   53
            Top             =   1020
            Width           =   1035
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
            TabIndex        =   52
            Top             =   330
            Width           =   1020
         End
      End
      Begin VB.Frame fraBody 
         BackColor       =   &H006B4B01&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1965
         Index           =   5
         Left            =   7350
         TabIndex        =   43
         Top             =   150
         Width           =   3885
         Begin VB.TextBox txtResultTag 
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
            MaxLength       =   15
            TabIndex        =   2
            Top             =   0
            Width           =   2565
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
            TabIndex        =   8
            Top             =   660
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
            TabIndex        =   22
            Top             =   1650
            Width           =   2325
         End
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
            MaxLength       =   50
            TabIndex        =   5
            Top             =   330
            Width           =   2565
         End
         Begin MSComCtl2.DTPicker dtpRequestDate 
            Height          =   285
            Left            =   1230
            TabIndex        =   18
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
            Format          =   301137921
            UpDown          =   -1  'True
            CurrentDate     =   38406
         End
         Begin MSComCtl2.DTPicker dtpRequestTime 
            Height          =   285
            Left            =   2550
            TabIndex        =   19
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
            Format          =   301137923
            UpDown          =   -1  'True
            CurrentDate     =   38406
         End
         Begin MSComCtl2.DTPicker dtpResultDate 
            Height          =   285
            Left            =   1230
            TabIndex        =   13
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
            Format          =   301137921
            UpDown          =   -1  'True
            CurrentDate     =   38406
         End
         Begin MSComCtl2.DTPicker dtpResultTime 
            Height          =   285
            Left            =   2550
            TabIndex        =   14
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
            Format          =   301137923
            UpDown          =   -1  'True
            CurrentDate     =   38406
         End
         Begin VB.Label lblResulttag 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FCE2C7&
            BackStyle       =   0  'Transparent
            Caption         =   "Stat/Routine:"
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
            Left            =   210
            TabIndex        =   72
            Top             =   60
            Width           =   990
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
            TabIndex        =   49
            Top             =   1650
            Width           =   285
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
            TabIndex        =   48
            Top             =   690
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
            TabIndex        =   47
            Top             =   1680
            Width           =   1215
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
            TabIndex        =   46
            Top             =   1020
            Width           =   1230
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
            TabIndex        =   45
            Top             =   1350
            Width           =   1230
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
            TabIndex        =   44
            Top             =   360
            Width           =   1215
         End
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5775
      Index           =   1
      Left            =   60
      TabIndex        =   36
      Top             =   2610
      Width           =   11295
      Begin VB.ComboBox cboResult 
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
         Height          =   330
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   1440
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton cmdStringValue 
         Caption         =   "Enter Result..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         TabIndex        =   75
         Top             =   1020
         Visible         =   0   'False
         Width           =   1725
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   150
         Top             =   5100
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmResultEntry.frx":0000
               Key             =   "remarks"
            EndProperty
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtRemarks 
         Height          =   585
         Left            =   2790
         TabIndex        =   26
         Top             =   5130
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   1032
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmResultEntry.frx":005E
      End
      Begin VB.TextBox txtEntry 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   24
         Top             =   540
         Visible         =   0   'False
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgResult 
         Height          =   4665
         Left            =   120
         TabIndex        =   25
         Top             =   150
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   8229
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   12615680
         ForeColorFixed  =   16777215
         FocusRect       =   0
         HighLight       =   0
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
      Begin RichTextLib.RichTextBox txtRemarks2 
         Height          =   585
         Left            =   6990
         TabIndex        =   78
         Top             =   5130
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   1032
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmResultEntry.frx":00DF
      End
      Begin VB.Label lblRemarks2 
         Appearance      =   0  'Flat
         BackColor       =   &H00A66B2B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remarks 2:"
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
         Height          =   270
         Left            =   6990
         TabIndex        =   74
         Top             =   4830
         Width           =   4185
      End
      Begin VB.Label lblDiff 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   240
         Left            =   2190
         TabIndex        =   42
         Top             =   4860
         Width           =   585
      End
      Begin VB.Label lblTextDiff 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Differential Count Total:"
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
         Height          =   240
         Left            =   120
         TabIndex        =   41
         Top             =   4860
         Width           =   2085
      End
      Begin VB.Label lblRemarks 
         Appearance      =   0  'Flat
         BackColor       =   &H00A66B2B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remarks:"
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
         Height          =   270
         Left            =   2790
         TabIndex        =   38
         Top             =   4830
         Width           =   4155
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5895
      Index           =   2
      Left            =   60
      TabIndex        =   39
      Top             =   2490
      Width           =   11265
      Begin VB.Shape Shape1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   3
         Height          =   465
         Left            =   3660
         Top             =   2400
         Width           =   3975
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please select a Profile."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   525
         Left            =   60
         TabIndex        =   40
         Top             =   2460
         Width           =   11205
      End
   End
   Begin VB.Frame fraBody 
      Appearance      =   0  'Flat
      BackColor       =   &H006B4B01&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Index           =   3
      Left            =   30
      TabIndex        =   37
      Top             =   8370
      Width           =   11295
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
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
         Left            =   1410
         TabIndex        =   28
         Top             =   180
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "Previous"
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
         Height          =   345
         Left            =   60
         TabIndex        =   27
         Top             =   180
         Width           =   1335
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
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
         Left            =   8190
         TabIndex        =   31
         Top             =   150
         Width           =   1545
      End
      Begin VB.CommandButton cmdSavePreview 
         Caption         =   "Save / Pre&view"
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
         Left            =   5100
         TabIndex        =   29
         Top             =   150
         Width           =   1455
      End
      Begin VB.CommandButton cmdSavePrint 
         Caption         =   "Save / Prin&t"
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
         Left            =   6600
         TabIndex        =   30
         Top             =   150
         Width           =   1545
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
         Height          =   405
         Left            =   9780
         TabIndex        =   32
         Top             =   150
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmResultEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private yInstrumentID As String
Private yPatientName As String
Private yResultDate As String
Private ySequenceNumber As String
Private yProfileID As String
Private yClassID As String

Private yTop As Integer
Private yLeft As Integer
Private yHeight As Integer
Private yWidth As Integer

Private tmpPatientName As String
Private tmpResultDate As String
Private tmpSequenceNumber As String
Private tmpArray() As String

Private yMultiStart As Long
Private yMultiTotal As Long
Private yMultiCurrent As Long

Private rstPatient As ADODB.Recordset
Private rstPatient2 As ADODB.Recordset
Private rstSavedResult As ADODB.Recordset
Private rstClass As ADODB.Recordset
Private rstAdditionalProfiles As ADODB.Recordset

Private yClassForms As New ClassForms

Private freeDelete As Boolean
Private HasRefNum As Boolean

Private AdditionalProfiles As String

Private marrPatientName() As String
Private marrResultDate() As String
Private marrSequenceNum() As String

Private NewVersion As Boolean
Private SIDSuffix As String
Private PIDSuffix As String
Private NewResult As Boolean
Private ShowPrintCopy As Boolean

Private Sub cboResult_Change()
    With msfgResult
        .TextMatrix(.Row, .Col) = cboResult.Text
    End With
End Sub

Private Sub cboResult_Click()
    cboResult_Change
End Sub

Private Sub cboResult_KeyDown(KeyCode As Integer, Shift As Integer)
With msfgResult
If KeyCode = vbKeyReturn Then
    If Shift = 1 Then
            If .Row = .Rows - 1 Then
                .Row = .Row
            Else
                .Row = .Row - 1
                msfgResult_EnterCell
            End If
    Else
            If .Row = .Rows - 1 Then
                .Row = .Row
                If txtRemarks.Visible Then txtRemarks.SetFocus
                txtEntry.Visible = False
            Else
                .Row = .Row + 1
                msfgResult_EnterCell
            End If
    End If
    highLight
End If
If KeyCode = vbKeyF3 Then
        .TextMatrix(.Row, 0) = yClassForms.showInputBox(.TextMatrix(.Row, 1) & " - String Value", , , .TextMatrix(.Row, 0), 50, , , True)
        If Len(Trim(.TextMatrix(.Row, 0))) > 0 Then
            .Col = 0
            .CellBackColor = .CellForeColor
        Else
            .Col = 0
            .CellBackColor = .BackColorFixed
        End If
        .Visible = False
        .Visible = True
End If
If KeyCode = vbKeyDown Then
    KeyCode = 0
    If .Row = .Rows - 1 Then
        .Row = .Row
'        If txtRemarks.Visible Then txtRemarks.SetFocus
'        txtEntry.Visible = False
    Else
        .Row = .Row + 1
        msfgResult_EnterCell
    End If
End If
If KeyCode = vbKeyUp Then
    KeyCode = 0
            If .Row = .Rows - 1 Then
                .Row = .Row
            Else
                .Row = .Row - 1
                msfgResult_EnterCell
            End If
End If
If KeyCode = vbKeyDelete Then
    cboResult.ListIndex = 0
End If

End With
End Sub

Private Sub cmdClose_Click()
    ReturnValueInt = CStr(yMultiTotal + 1)
    Unload Me
End Sub

Public Property Let wTop(ByVal vNewValue As Integer)
    yTop = vNewValue
End Property

Public Property Let wLeft(ByVal vNewValue As Integer)
    yLeft = vNewValue
End Property

Public Property Let wHeight(ByVal vNewValue As Integer)
    yHeight = vNewValue
End Property

Public Property Let wWidth(ByVal vNewValue As Integer)
    yWidth = vNewValue
End Property

Public Property Let wInstrumentID(ByVal vNewValue As String)
    yInstrumentID = vNewValue
End Property

Public Property Let wPatientName(ByVal vNewValue As String)
    yPatientName = vNewValue
End Property

Public Property Let wResultdate(ByVal vNewValue As String)
    yResultDate = vNewValue
End Property

Public Property Let wSequenceNumber(ByVal vNewValue As String)
    ySequenceNumber = vNewValue
End Property


Private Sub cmdNext_Click()
Me.MousePointer = 11
If NewVersion Then
    yMultiCurrent = yMultiCurrent + 1
    ClearValues
    InitForm
    If yMultiTotal = yMultiCurrent Then cmdNext.Enabled = False Else cmdNext.Enabled = True
Else
    If yMultiTotal <> yMultiCurrent Then ReturnValueInt = yMultiCurrent + 1
    Unload Me
End If
Me.MousePointer = 0
End Sub

Private Sub cmdPrevious_Click()
Me.MousePointer = 11
If NewVersion Then
    yMultiCurrent = yMultiCurrent - 1
    ClearValues
    InitForm
    If yMultiCurrent = 1 Then cmdPrevious.Enabled = False Else cmdPrevious.Enabled = True
Else
    If yMultiStart <> yMultiCurrent Then ReturnValueInt = yMultiCurrent - 1
    Unload Me
End If
Me.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
Dim willNotValidate As Boolean
If Len(Trim(lblClassName.Caption)) > 0 And Len(Trim(lblProfileName.Caption)) > 0 And fraBody(1).Visible = True And Len(Trim(txtLastName.Text)) > 0 Then
    If NewResult Then
        If Not HasRefNum Then saveNewResult
    End If
    If HasRefNum Then
        SaveResult
        willNotValidate = False
    Else
        SaveResult False, True
        willNotValidate = True
    End If
'    MsgBox getUpdatedValues(txtRefNum.Text)
    'If willNotValidate Or validateDiff Or Val(GetSetting("HostForms_SA", "Settings", "ByPassDiff", "0")) = 1 Then
     If willNotValidate Or Val(GetSetting("HostForms_SA", "Settings", "ByPassDiff", "0")) = 1 Then
        If (yMultiCurrent - yMultiStart) + 1 = (yMultiTotal - yMultiStart) + 1 Then
            ReturnValueInt = CStr(yMultiTotal + 1)
            Unload Me
        Else
            If cmdNext.Visible Then
                cmdNext_Click
            End If
        End If
    End If
Else
    MsgBox "Cannot Save Result... Please Enter First Name and Last Name.", vbOKOnly + vbInformation, "Result Entry"
End If
End Sub

Private Sub cmdSavePreview_Click()
If Len(Trim(lblClassName.Caption)) > 0 And Len(Trim(lblProfileName.Caption)) > 0 And fraBody(1).Visible = True And Len(Trim(txtLastName.Text)) > 0 Then
    If NewResult Then
        If Not HasRefNum Then saveNewResult
    End If
    SaveResult
    If validateDiff Or Val(GetSetting("HostForms_SA", "Settings", "ByPassDiff", "0")) = 1 Then
        Me.MousePointer = 11
        If OpenMainReport(App.Path & "\reports\" & getDefaultTemplate(txtRefNum.Text), txtRefNum.Text & "", "") <> "Error" Then ShowReportViewer False, "Result Print Preview", False, txtRefNum.Text
        Me.MousePointer = 0
    End If
Else
    MsgBox "Cannot Save Result...", vbOKOnly + vbInformation, "Result Entry"
End If
'Me.Refresh
Me.fraTitle.Visible = False
Me.fraTitle.Visible = True
Me.fraBody(0).Visible = False
Me.fraBody(0).Visible = True
Me.fraBody(1).Visible = False
Me.fraBody(1).Visible = True
Me.fraBody(3).Visible = False
Me.fraBody(3).Visible = True
End Sub

Private Sub saveNewResult()
Dim temps As String
If yPatientName & yResultDate & ySequenceNumber = "NEWNEWNEW" Then
    If DBServerType = "MSSQL SERVER" Then
        temps = "Insert into tbLabInstrumentResults (InstrumentID, Resultdate, Accessionnum, PatientName, Status, ProfileExam, AccessionnumOrg, IdNum, BirthDate, NewPatientName,Printed,Refnum, ResultID)" & vbCrLf & _
            "Select '" & InstrumentID & "','" & Format(dtpResultDate.Value, "mm/dd/yyyy") & " " & Format(dtpResultTime.Value, "HH:mm") & "','" & txtSpecimenID.Text & "','" & txtLastName.Text & ", " & txtFirstName.Text & " " & txtMiddleName.Text & "','3','" & txtProfileID.Text & "','" & txtSpecimenID.Text & "','" & txtPatientID.Text & "','" & _
            txtBirthDate.Text & "','" & txtLastName.Text & ", " & txtFirstName.Text & " " & txtMiddleName.Text & "'," & "0" & ",'" & txtRefNum.Text & "','" & GenerateResultID(txtRefNum.Text) & "';"
    Else
        temps = "Insert into tbLabInstrumentResults (InstrumentID, Resultdate, Accessionnum, PatientName, Status, ProfileExam, AccessionnumOrg, IdNum, BirthDate, NewPatientName,Printed,Refnum, ResultID)" & vbCrLf & _
            "Select '" & InstrumentID & "',#" & Format(dtpResultDate.Value, "mm/dd/yyyy") & " " & Format(dtpResultTime.Value, "HH:mm") & "#,'" & txtSpecimenID.Text & "','" & txtLastName.Text & ", " & txtFirstName.Text & " " & txtMiddleName.Text & "','3','" & txtProfileID.Text & "','" & txtSpecimenID.Text & "','" & txtPatientID.Text & "','" & _
            txtBirthDate.Text & "','" & txtLastName.Text & ", " & txtFirstName.Text & " " & txtMiddleName.Text & "'," & False & ",'" & txtRefNum.Text & "','" & GenerateResultID(txtRefNum.Text) & "';"
    End If
    myConnection.Execute temps
    temps = "Insert into tbLabInstrumentRValues (ResultID) Select '" & GenerateResultID(txtRefNum.Text) & "'"
    myConnection.Execute temps
End If
End Sub

Private Sub cmdStringValue_Click()
Dim tempStr As String
With msfgResult
        tempStr = yClassForms.showInputBox(Trim(.TextMatrix(.Row, 1)) & " - String Value", , , .TextMatrix(.Row, 0), 254, , , True)
        If Len(Trim(tempStr)) > 0 Then
            .TextMatrix(.Row, 0) = tempStr
        End If
        If Len(Trim(.TextMatrix(.Row, 0))) > 0 Then
            .Col = 0
            .CellBackColor = .CellForeColor
            .Col = 3
'            .CellBackColor = &H8000000F
'            .CellForeColor = &H8000000F
            .CellForeColor = vbButtonText
            .TextMatrix(.Row, 3) = IIf(Len(Trim(.TextMatrix(.Row, 0))) > 12, Mid(.TextMatrix(.Row, 0), 1, 12) & "...", .TextMatrix(.Row, 0))
        Else
            .Col = 0
            .CellBackColor = .BackColorFixed
'            .CellBackColor = vbWindowBackground
'            .CellForeColor = vbButtonText
            .TextMatrix(.Row, 3) = ""
        End If
        .Visible = False
        .Visible = True
End With
End Sub

Private Sub dtpRequestDate_GotFocus()
    dtpRequestDate.CalendarBackColor = &HC0E0FF
End Sub

Private Sub dtpRequestDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub dtpRequestDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub dtpRequestDate_LostFocus()
    dtpResultDate.CalendarBackColor = &H80000005
End Sub

Private Sub dtpRequestTime_GotFocus()
    dtpRequestTime.CalendarBackColor = &HC0E0FF
End Sub

Private Sub dtpRequestTime_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub dtpRequestTime_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub dtpRequestTime_LostFocus()
    dtpResultDate.CalendarBackColor = &H80000005
End Sub

Private Sub dtpResultDate_GotFocus()
    dtpResultDate.CalendarBackColor = &HC0E0FF
End Sub

Private Sub dtpResultDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub dtpResultDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub dtpResultDate_LostFocus()
    dtpResultDate.CalendarBackColor = &HC0E0FF
End Sub

Private Sub dtpResultTime_GotFocus()
    dtpResultTime.CalendarBackColor = &H80000005
End Sub

Private Sub dtpResultTime_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub dtpResultTime_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub dtpResultTime_LostFocus()
    dtpResultDate.CalendarBackColor = &H80000005
End Sub

Private Sub Form_Load()
setFormColors Me
lblText(0).ForeColor = lblName.ForeColor
If Val(GetSetting("HostForms_SA", "Settings", "AutoResize", "0")) = 1 Then
    Me.Top = yTop
    Me.Left = yLeft
    Me.Height = yHeight
    Me.Width = yWidth
End If
lblRemarks.BackColor = fraTitle.BackColor
lblRemarks.ForeColor = lblTitle.ForeColor
txtPatientID.Enabled = Val(GetSetting("HostForms_SA", "Settings", "PatientID", "1"))
txtSpecimenID.Enabled = Val(GetSetting("HostForms_SA", "Settings", "AccNum", "1"))
txtAge.Enabled = Val(GetSetting("HostForms_SA", "Settings", "Age", "1"))
txtSex.Enabled = Val(GetSetting("HostForms_SA", "Settings", "Sex", "1"))
txtWard.Enabled = Val(GetSetting("HostForms_SA", "Settings", "Ward", "1"))
txtRoom.Enabled = Val(GetSetting("HostForms_SA", "Settings", "Room", "1"))
dtpResultDate.Enabled = Val(GetSetting("HostForms_SA", "Settings", "ResultDate", "1"))
dtpResultTime.Enabled = Val(GetSetting("HostForms_SA", "Settings", "ResultDate", "1"))
dtpRequestDate.Enabled = Val(GetSetting("HostForms_SA", "Settings", "RequestDate", "1"))
dtpRequestTime.Enabled = Val(GetSetting("HostForms_SA", "Settings", "RequestDate", "1"))
txtPhysician.Enabled = Val(GetSetting("HostForms_SA", "Settings", "Physician", "1"))
txtMedtech.Enabled = Val(GetSetting("HostForms_SA", "Settings", "MedTech", "1"))
txtBirthDate.Enabled = Val(GetSetting("HostForms_SA", "Settings", "BirthDate", "1"))
txtCivilStatus.Enabled = Val(GetSetting("HostForms_SA", "Settings", "CivilStatus", "1"))
txtSpecimen.Enabled = Val(GetSetting("HostForms_SA", "Settings", "Specimen", "1"))
txtMiddleName.Enabled = Val(GetSetting("HostForms_SA", "Settings", "MiddleName", "0"))
txtResultTag.Enabled = Val(GetSetting("HostForms_SA", "Settings", "RoutineStat", "1"))
ShowPrintCopy = Val(GetSetting("HostForms_SA", "Settings", "ShowPrintCopy", "0"))
If txtPatientID.Enabled Then txtPatientID.BackColor = vbWindowBackground Else txtPatientID.BackColor = vbButtonFace
If txtSpecimenID.Enabled Then txtSpecimenID.BackColor = vbWindowBackground Else txtSpecimenID.BackColor = vbButtonFace
If txtAge.Enabled Then txtAge.BackColor = vbWindowBackground Else txtAge.BackColor = vbButtonFace
If txtSex.Enabled Then txtSex.BackColor = vbWindowBackground Else txtSex.BackColor = vbButtonFace
If txtWard.Enabled Then txtWard.BackColor = vbWindowBackground Else txtWard.BackColor = vbButtonFace
If txtRoom.Enabled Then txtRoom.BackColor = vbWindowBackground Else txtRoom.BackColor = vbButtonFace
If txtPhysician.Enabled Then txtPhysician.BackColor = vbWindowBackground Else txtPhysician.BackColor = vbButtonFace
If txtMedtech.Enabled Then txtMedtech.BackColor = vbWindowBackground Else txtMedtech.BackColor = vbButtonFace
If txtBirthDate.Enabled Then txtBirthDate.BackColor = vbWindowBackground Else txtBirthDate.BackColor = vbButtonFace
If txtCivilStatus.Enabled Then txtCivilStatus.BackColor = vbWindowBackground Else txtCivilStatus.BackColor = vbButtonFace
If txtSpecimen.Enabled Then txtSpecimen.BackColor = vbWindowBackground Else txtSpecimen.BackColor = vbButtonFace
If txtMiddleName.Enabled Then txtMiddleName.BackColor = vbWindowBackground Else txtMiddleName.BackColor = vbButtonFace
If txtResultTag.Enabled Then txtResultTag.BackColor = vbWindowBackground Else txtResultTag.BackColor = vbButtonFace

lblPatientID.Caption = GetSetting("HostForms_SA", "LabelSettings", "PatientID", "Patient ID")
lblSpecimenID.Caption = GetSetting("HostForms_SA", "LabelSettings", "SpecimenID", "Specimen ID")
lblAge.Caption = GetSetting("HostForms_SA", "LabelSettings", "Age", "Age")
lblSex.Caption = GetSetting("HostForms_SA", "LabelSettings", "Sex", "Sex")
lblWard.Caption = GetSetting("HostForms_SA", "LabelSettings", "Ward", "Ward")
lblRoom.Caption = GetSetting("HostForms_SA", "LabelSettings", "Room", "Room")
lblResultdate.Caption = GetSetting("HostForms_SA", "LabelSettings", "Resultdate", "Result Date")
lblRequestDate.Caption = GetSetting("HostForms_SA", "LabelSettings", "RequestDate", "Request Date")
lblPhysician.Caption = GetSetting("HostForms_SA", "LabelSettings", "Physician", "Physician")
lblMedtech.Caption = GetSetting("HostForms_SA", "LabelSettings", "Medtech", "Med. Tech.")
lblRemarks.Caption = GetSetting("HostForms_SA", "LabelSettings", "Remarks", "Remarks")
lblRemarks2.Caption = GetSetting("HostForms_SA", "LabelSettings", "Remarks2", "Remarks 2")
lblBirthDate.Caption = GetSetting("HostForms_SA", "LabelSettings", "BirthDate", "Birth Date")
lblCivilStatus.Caption = GetSetting("HostForms_SA", "LabelSettings", "CivilStatus", "Civil Status")
lblSpecimen.Caption = GetSetting("HostForms_SA", "LabelSettings", "Specimen", "Specimen")
lblResulttag.Caption = GetSetting("HostForms_SA", "LabelSettings", "RoutineStat", "Routine/Stat")

Select Case Val(GetSetting("HostForms_SA", "Settings", "Remarks", "1")) & Val(GetSetting("HostForms_SA", "Settings", "Remarks2", "0"))
    Case "00"
        lblRemarks.Visible = False
        txtRemarks.Visible = False
        lblRemarks2.Visible = False
        txtRemarks2.Visible = False
    Case "01"
        lblRemarks.Visible = False
        txtRemarks.Visible = False
        lblRemarks2.Visible = True
        txtRemarks2.Visible = True
        lblRemarks2.Width = 8385
        txtRemarks2.Width = 8385
        lblRemarks2.Left = 2790
        txtRemarks2.Left = 2790
    Case "10"
        lblRemarks.Visible = True
        txtRemarks.Visible = True
        lblRemarks2.Visible = False
        txtRemarks2.Visible = False
        lblRemarks.Width = 8385
        txtRemarks.Width = 8385
    Case "11"
        lblRemarks.Visible = True
        txtRemarks.Visible = True
        lblRemarks2.Visible = True
        txtRemarks2.Visible = True
End Select

If SectionCode = "PULMO" Then
    lblText(6).Caption = "Pulmonologist"
Else
    lblText(6).Caption = "Pathologist"
End If

lblPatientID.ForeColor = lblText(1).ForeColor
lblSpecimenID.ForeColor = lblText(1).ForeColor
lblAge.ForeColor = lblText(1).ForeColor
lblSex.ForeColor = lblText(1).ForeColor
lblWard.ForeColor = lblText(1).ForeColor
lblRoom.ForeColor = lblText(1).ForeColor
lblResultdate.ForeColor = lblText(1).ForeColor
lblRequestDate.ForeColor = lblText(1).ForeColor
lblPhysician.ForeColor = lblText(1).ForeColor
lblMedtech.ForeColor = lblText(1).ForeColor
'lblRemarks.ForeColor = lblText(1).ForeColor
lblBirthDate.ForeColor = lblText(1).ForeColor
lblCivilStatus.ForeColor = lblText(1).ForeColor
lblSpecimen.ForeColor = lblText(1).ForeColor
lblResulttag.ForeColor = lblText(1).ForeColor

txtWard.Locked = GetSetting("HostForms_SA", "SelList", "cWard", "0") = 1
txtRoom.Locked = GetSetting("HostForms_SA", "SelList", "cRoom", "0") = 1
txtCivilStatus.Locked = GetSetting("HostForms_SA", "SelList", "cCivil Status", "0") = 1
txtSpecimen.Locked = GetSetting("HostForms_SA", "SelList", "cSpecimen", "0") = 1
txtResultTag.Locked = GetSetting("HostForms_SA", "SelList", "cRoutineStat", "0") = 1

SIDSuffix = GetSetting("HostForms_SA", "Settings", "SIDSuffix", "")
NewResult = False
'If Val(GetSetting("HostForms_SA", "Settings", "Remarks", "1")) = 1 Then
'    lblRemarks.Visible = True
'    txtRemarks.Visible = True
'Else
'    lblRemarks.Visible = False
'    txtRemarks.Visible = False
'End If
TotalDiff = GetSetting("LAB-i", "Settings", "DiffCount", "1")

NewVersion = False
If InStr(1, yResultDate, "|") > 0 Then
    marrPatientName = Split(yPatientName, "|")
    marrResultDate = Split(yResultDate, "|")
    marrSequenceNum = Split(ySequenceNumber, "|")
    tmpPatientName = yPatientName
    tmpResultDate = yResultDate
    tmpSequenceNumber = ySequenceNumber
    yPatientName = marrPatientName(0)
    yResultDate = marrResultDate(0)
    ySequenceNumber = marrSequenceNum(0)
    yMultiStart = 1
    yMultiTotal = UBound(marrResultDate) + 1
    yMultiCurrent = 1
    NewVersion = True
End If

ReDim tmpArray((yMultiTotal - yMultiStart) + 1) As String
InitForm

Select Case GetSetting("HostForms_SA", "Settings", "DefaultFieldPos", "Patient ID")
    Case "Patient ID"
        If txtPatientID.Visible Then txtPatientID.SetFocus
    Case "Sample ID"
        If txtSpecimenID.Visible Then txtSpecimenID.SetFocus
    Case "Last Name"
        If txtLastName.Visible Then txtLastName.SetFocus
    Case "First Name"
        If txtFirstName.Visible Then txtFirstName.SetFocus
End Select
'Call WheelHook(Me.hWnd)
End Sub

Private Sub InitForm()
Dim rstTemp As ADODB.Recordset
Dim ctr As Integer
If NewVersion Then
    yPatientName = marrPatientName(yMultiCurrent - 1)
    yResultDate = marrResultDate(yMultiCurrent - 1)
    ySequenceNumber = marrSequenceNum(yMultiCurrent - 1)
End If
If yPatientName & yResultDate & ySequenceNumber = "NEWNEWNEW" Then
    txtRefNum.Text = Format(Date, "mmddyyyy") & Format(Time, "hhmmss") & "N"
    HasRefNum = False
    cmdSavePreview.Enabled = False
    cmdSavePrint.Enabled = False
    cmdSave.Enabled = False
    msfgResult.Cols = 23
    dtpRequestDate.Value = Format(Now, "mm/dd/yyyy hh:mm")
    dtpRequestTime.Value = Format(Now, "mm/dd/yyyy hh:mm")
    dtpResultDate.Value = Format(Now, "mm/dd/yyyy")
    dtpResultTime.Value = Format(Now, "mm/dd/yyyy hh:mm")
    txtProfileID.Text = yProfileID
    If DBServerType = "LOCAL DATABASE" Then
        Set rstTemp = myConnection.Execute("Select * from tblabpathologists where Default = True")
    ElseIf DBServerType = "MSSQL SERVER" Then
        Set rstTemp = myConnection.Execute("Select * from tblabpathologists where [Default] = 1")
    End If
    If rstTemp.RecordCount > 0 Then
        txtPathologist.Text = rstTemp.Fields("Pathologist").Value & ""
    Else
        txtPathologist.Text = ""
    End If
    If Val(GetSetting("HostForms_SA", "Settings", "SaveDefaultMT", "0")) = 1 Then
        txtMedtech.Text = GetSetting("HostForms_SA", "Settings", "SaveDefaultMTName", "")
    End If
    DisplayResultForm False
    NewResult = True
Else
    If DBServerType = "LOCAL DATABASE" Then
        If Trim(yPatientName) = "" Then
            Set rstPatient = myConnection.Execute("sp_getInstrumentResultDemo '" & yInstrumentID & "',' ',#" & Format(yResultDate, "mm/dd/yyyy HH:mm:ss") & "#")
        Else
            Set rstPatient = myConnection.Execute("sp_getInstrumentResultDemo '" & yInstrumentID & "','" & Replace(yPatientName, "'", "''") & "',#" & Format(yResultDate, "mm/dd/yyyy HH:mm:ss") & "#")
        End If
    ElseIf DBServerType = "MSSQL SERVER" Then
        If Trim(yPatientName) = "" Then
            Set rstPatient = myConnection.Execute("sp_getInstrumentResultDemo '" & yInstrumentID & "',' ','" & Format(yResultDate, "mm/dd/yyyy HH:mm:ss") & "'")
        Else
            Set rstPatient = myConnection.Execute("sp_getInstrumentResultDemo '" & yInstrumentID & "','" & Replace(yPatientName, "'", "''") & "','" & Format(yResultDate, "mm/dd/yyyy HH:mm:ss") & "'")
        End If
    End If
    rstPatient.MoveFirst
    setValues
End If
If Right(txtRefNum.Text, 1) = "N" Then NewResult = True
setGrid
fillLists
If yMultiStart = yMultiTotal Then
    cmdNext.Visible = False
    cmdPrevious.Visible = False
    lblTitle.Caption = "Result Entry and Verification"
ElseIf yMultiStart = yMultiCurrent Then
    cmdNext.Enabled = True
    cmdPrevious.Enabled = False
ElseIf yMultiTotal = yMultiCurrent Then
    cmdNext.Enabled = False
    cmdPrevious.Enabled = True
Else
    cmdNext.Enabled = True
    cmdPrevious.Enabled = True
End If
If HasRefNum Then
    lblTitle.Caption = "Result Editing and Reprinting - [" & (yMultiCurrent - yMultiStart) + 1 & " of " & (yMultiTotal - yMultiStart) + 1 & "]"
Else
    If NewResult Then
        lblTitle.Caption = "Result Entry and Verification - [NEW RECORD]"
    Else
        lblTitle.Caption = "Result Entry and Verification - [" & (yMultiCurrent - yMultiStart) + 1 & " of " & (yMultiTotal - yMultiStart) + 1 & "]"
    End If
End If

For ctr = 1 To (yMultiTotal - yMultiStart) + 1
    tmpArray(ctr) = ""
Next ctr

End Sub

Private Sub ClearValues()
    txtRefNum.Text = Format(Date, "mmddyyyy") & Format(Time, "hhmmss")
    txtLastName.Text = ""
    txtFirstName.Text = ""
    txtMiddleName.Text = ""
    txtAge.Text = ""
    txtSex.Text = ""
    txtRoom.Text = ""
    txtWard.Text = ""
'    txtProfileID.Text = ""
    txtPhysician.Text = ""
    txtMedtech.Text = ""
    txtPatientID.Text = ""
    txtSpecimenID.Text = ""
    txtRemarks.Text = ""
    txtPathologist.Text = ""
    dtpRequestDate.Value = Format(Date, "mm/dd/yyyy hh:mm")
    dtpRequestTime.Value = Format(Date, "mm/dd/yyyy hh:mm")
    dtpResultDate.Value = Format(Date, "mm/dd/yyyy")
    dtpResultTime.Value = Format(Date, "mm/dd/yyyy hh:mm")
    AdditionalProfiles = ""
    txtClassID.Text = ""
End Sub

Private Sub clearValuesN()
'    txtRefNum.Text = Format(Date, "mmddyyyy") & Format(Time, "hhmmss") & "N"
    txtLastName.Text = ""
    txtFirstName.Text = ""
    txtMiddleName.Text = ""
    txtBirthDate.Text = ""
    txtCivilStatus.Text = ""
    txtSpecimen.Text = ""
    txtAge.Text = ""
    txtSex.Text = ""
    txtRoom.Text = ""
    txtWard.Text = ""
    txtPhysician.Text = ""
    txtMedtech.Text = ""
    txtSpecimenID.Text = ""
    txtRemarks.Text = ""
    txtPathologist.Text = ""
    dtpRequestDate.Value = Format(Now, "mm/dd/yyyy HH:mm")
    dtpRequestTime.Value = Format(Now, "mm/dd/yyyy HH:mm")
    dtpResultDate.Value = Format(Now, "mm/dd/yyyy HH:mm")
    dtpResultTime.Value = Format(Now, "mm/dd/yyyy HH:mm")
    AdditionalProfiles = ""
    txtClassID.Text = ""
End Sub

Private Sub setValues()
Dim tempPatientName As String
Dim Temp() As String
Dim temp2() As String
Dim rstTemp As ADODB.Recordset
On Error GoTo errFound

If Len(Trim(IIf(IsNull(rstPatient.Fields("RefNum").Value), "", rstPatient.Fields("RefNum").Value))) > 0 Then
    txtRefNum.Text = rstPatient.Fields("RefNum").Value
    HasRefNum = True
    Set rstPatient2 = myConnection.Execute("Select * from tbPatientMaster where refnum = '" & txtRefNum & "'")
    If rstPatient2.RecordCount > 0 Then
        txtLastName.Text = rstPatient2.Fields("LastName").Value
        txtFirstName.Text = rstPatient2.Fields("FirstName").Value
        txtMiddleName.Text = rstPatient2.Fields("MiddleName").Value
        txtAge.Text = rstPatient2.Fields("Age").Value
        txtSex.Text = rstPatient2.Fields("Sex").Value
        txtRoom.Text = rstPatient2.Fields("Room").Value
        txtWard.Text = rstPatient2.Fields("Ward").Value
        txtProfileID.Text = rstPatient2.Fields("ProfileID").Value
        If Len(Trim(rstPatient2.Fields("ProfileID").Value)) > 0 Then lblProfileName.Caption = yClassForms.showSearchCode(1, txtProfileID.Text, True)
        txtPhysician.Text = rstPatient2.Fields("Physician").Value
        txtMedtech.Text = rstPatient2.Fields("MedTech").Value
        txtPatientID.Text = rstPatient2.Fields("PatientID").Value & ""
        txtSpecimenID.Text = rstPatient2.Fields("SpecimenID").Value & ""
        txtRemarks.Text = rstPatient2.Fields("Remarks").Value & ""
        txtRemarks2.Text = rstPatient2.Fields("Remarks2").Value & ""
        txtPathologist.Text = rstPatient2.Fields("Pathologist").Value & ""
        dtpRequestDate.Value = Format(rstPatient2.Fields("RequestDate").Value, "mm/dd/yyyy hh:mm")
        dtpRequestTime.Value = Format(rstPatient2.Fields("RequestDate").Value, "mm/dd/yyyy hh:mm")
        dtpResultDate.Value = Format(rstPatient2.Fields("ResultDate").Value, "mm/dd/yyyy")
        dtpResultTime.Value = Format(rstPatient2.Fields("ResultDate").Value, "mm/dd/yyyy hh:mm")
        AdditionalProfiles = rstPatient2.Fields("AdditionalProfiles").Value & ""
        txtBirthDate.Text = rstPatient2.Fields("BirthDate").Value & ""
        txtCivilStatus.Text = rstPatient2.Fields("CivilStatus").Value & ""
        txtSpecimen.Text = rstPatient2.Fields("Specimen").Value & ""
        txtResultTag.Text = rstPatient2.Fields("ResultFlag").Value & ""
        If Len(Trim(AdditionalProfiles)) > 0 And txtProfileID = "0" Then lblProfileName.Caption = "Others(" & AdditionalProfiles & ")"
    Else
        MsgBox "Patient Information not found. Press OK to continue.", vbInformation + vbOKOnly, App.Title
        txtRefNum.Text = rstPatient.Fields("RefNum").Value
        HasRefNum = False
        rstPatient.MoveFirst
        tempPatientName = Replace(rstPatient.Fields("PatientName").Value, ",,", ",")
        If Val(GetSetting("HostForms_SA", "Settings", "CheckID", "0")) = 1 Then
            temp2 = Split(GetIdFromName(tempPatientName), "|")
            tempPatientName = temp2(1)
        End If
        If InStr(1, tempPatientName, ",") > 0 Then
            Temp = Split(tempPatientName, ",")
            txtLastName.Text = Trim(Temp(0))
            txtFirstName.Text = Trim(Temp(1))
        ElseIf InStr(1, tempPatientName, " ") > 0 Then
            Temp = Split(tempPatientName, " ")
            txtLastName.Text = Trim(Temp(1))
            txtFirstName.Text = Trim(Temp(0))
        Else
            txtLastName.Text = tempPatientName
        End If
        If Val(GetSetting("HostForms_SA", "Settings", "CheckID", "0")) = 1 Then
            If Len(Trim(rstPatient.Fields("accessionnum").Value)) = 0 Then
                txtSpecimenID.Text = temp2(0)
            Else
                txtSpecimenID.Text = rstPatient.Fields("accessionnum").Value & SIDSuffix
            End If
        Else
            txtSpecimenID.Text = rstPatient.Fields("accessionnum").Value & SIDSuffix
        End If
        dtpRequestDate.Value = Format(yResultDate, "mm/dd/yyyy hh:mm")
        dtpRequestTime.Value = Format(yResultDate, "mm/dd/yyyy hh:mm")
        dtpResultDate.Value = Format(yResultDate, "mm/dd/yyyy")
        dtpResultTime.Value = Format(yResultDate, "mm/dd/yyyy hh:mm")
        txtProfileID.Text = GetSetting("LAB-i", "Settings", "ProfileID", "1")
        If Len(Trim(rstPatient.Fields("ProfileID").Value)) > 0 Then lblProfileName.Caption = yClassForms.showSearchCode(1, txtProfileID.Text)
'        If txtFirstName.Visible Then txtFirstName.SetFocus
    End If
Else
    txtRefNum.Text = Format(Date, "mmddyyyy") & Format(Time, "hhmmss")
    HasRefNum = False
    rstPatient.MoveFirst
    If Val(GetSetting("HostForms_SA", "Settings", "UseNewName", "0")) = 1 Then
        If Len(Trim(rstPatient.Fields("NewPatientName").Value & "")) > 0 Then
            tempPatientName = Replace(rstPatient.Fields("NewPatientName").Value, ",,", ",")
        Else
            tempPatientName = Replace(rstPatient.Fields("PatientName").Value, ",,", ",")
        End If
    Else
        tempPatientName = Replace(rstPatient.Fields("PatientName").Value, ",,", ",")
    End If
    If Val(GetSetting("HostForms_SA", "Settings", "CheckID", "0")) = 1 Then
        temp2 = Split(GetIdFromName(tempPatientName), "|")
        tempPatientName = temp2(1)
    End If
    If InStr(1, tempPatientName, ",") > 0 Then
        Temp = Split(tempPatientName, ",")
        txtLastName.Text = Trim(Temp(0))
        txtFirstName.Text = Trim(Temp(1))
    ElseIf InStr(1, tempPatientName, " ") > 0 Then
        Temp = Split(tempPatientName, " ")
        txtLastName.Text = Trim(Temp(1))
        txtFirstName.Text = Trim(Temp(0))
    Else
        txtLastName.Text = tempPatientName
    End If
    txtPatientID.Text = rstPatient.Fields("IDnum").Value & PIDSuffix
    If Val(GetSetting("HostForms_SA", "Settings", "CheckID", "0")) = 1 Then
        If Len(Trim(rstPatient.Fields("accessionnum").Value)) = 0 Then
            txtSpecimenID.Text = temp2(0)
        Else
            txtSpecimenID.Text = rstPatient.Fields("accessionnum").Value & SIDSuffix
        End If
    Else
        txtSpecimenID.Text = rstPatient.Fields("accessionnum").Value & SIDSuffix
    End If
    dtpRequestDate.Value = Format(yResultDate, "mm/dd/yyyy hh:mm")
    dtpRequestTime.Value = Format(yResultDate, "mm/dd/yyyy hh:mm")
    dtpResultDate.Value = Format(yResultDate, "mm/dd/yyyy hh:mm")
    dtpResultTime.Value = Format(yResultDate, "mm/dd/yyyy hh:mm")
    txtProfileID.Text = GetSetting("LAB-i", "Settings", "ProfileID", "1")
    txtBirthDate = Format(rstPatient.Fields("BirthDate").Value & "", ">")
    txtRoom.Text = rstPatient.Fields("Room").Value & ""
    txtAge.Text = rstPatient.Fields("Age").Value & ""
    txtSex.Text = rstPatient.Fields("Sex").Value & ""
    txtRemarks.Text = rstPatient.Fields("Comments").Value & ""
    txtCivilStatus.Text = rstPatient.Fields("Operator").Value & ""
    If txtSex.Text = "M" Or txtSex.Text = "MALE" Then
        txtSex.Text = "MALE"
    ElseIf txtSex.Text = "F" Or txtSex.Text = "FEMALE" Then
        txtSex.Text = "FEMALE"
    Else
        txtSex.Text = ""
    End If
    
    If Val(GetSetting("HostForms_SA", "Settings", "DetectNV", "0")) = 1 Then
        If Len(Trim(txtSex.Text)) > 0 Then
            txtClassID.Text = getNVclass(txtSex.Text)
        End If
    Else
        txtClassID.Text = GetSetting("LAB-i", "Settings", "ClassID", "")
    End If
    If Len(Trim(txtClassID.Text)) > 0 Then txtClassID_KeyDown vbKeyReturn, 0
    txtSpecimen.Text = rstPatient.Fields("Specimen").Value & ""
    txtPhysician.Text = rstPatient.Fields("Doctor").Value & ""
    If Val(GetSetting("HostForms_SA", "Settings", "SaveDefaultMT", "0")) = 1 Then
        txtMedtech.Text = GetSetting("HostForms_SA", "Settings", "SaveDefaultMTName", "")
    End If
    txtBirthDate_Change
    txtProfileID.Text = GetSetting("LAB-i", "Settings", "ProfileID", "1")
    If Len(txtProfileID.Text) > 0 Then lblProfileName.Caption = yClassForms.showSearchCode(1, txtProfileID.Text, False)
'    If Len(Trim(rstPatient.Fields("ProfileID").Value)) > 0 Then lblProfileName.Caption = yClassForms.showSearchCode(1, txtProfileID.Text, False)
'    If txtFirstName.Visible Then txtFirstName.SetFocus
    If DBServerType = "LOCAL DATABASE" Then
        Set rstTemp = myConnection.Execute("Select * from tblabpathologists where Default = True")
    Else
        Set rstTemp = myConnection.Execute("Select * from tblabpathologists where [Default] = 1")
    End If
    If rstTemp.RecordCount > 0 Then
        txtPathologist.Text = rstTemp.Fields("Pathologist").Value & ""
    Else
        txtPathologist.Text = ""
    End If
    '---------------------------------------
    'insert code here

    If Val(GetSetting("HostForms_SA", "Settings", "Search OBI", "0")) = 1 And Val(GetSetting("LAB-i", "OBISettings", "UseExDatabase", "0")) = 1 Then
        If GetSetting("LAB-i", "OBISettings", "ServerType", "") = "MS SQL Server 2000" Then
            Set rstTemp = myCon.Execute("Select TOP 1 * from wh_procre where PATIENT_SAMPLE_ID='txtSpecimenID.Text'")
        ElseIf GetSetting("LAB-i", "OBISettings", "ServerType", "") = "MySQL" Then
            Set rstTemp = myCon.Execute("Select * from Table where `PATIENT SAMPLE ID` = <FLD_1> '" & Replace(txtSpecimenID.Text, "''", "<>") & "' LIMIT 1")
        End If
        
        If rstTemp.RecordCount > 0 Then
            txtLastName.Text = rstTemp.Fields(GetSetting("LAB-i", "OBISettings", "FLD_3", "''")).Value & ""
            txtFirstName.Text = rstTemp.Fields(GetSetting("LAB-i", "OBISettings", "FLD_4", "''")).Value & ""
            txtMiddleName.Text = rstTemp.Fields(GetSetting("LAB-i", "OBISettings", "FLD_5", "''")).Value & ""
            txtBirthDate.Text = rstTemp.Fields(GetSetting("LAB-i", "OBISettings", "FLD_17", "''")).Value & ""
            txtAge.Text = rstTemp.Fields(GetSetting("LAB-i", "OBISettings", "FLD_6", "''")).Value & ""
            txtSex.Text = rstTemp.Fields(GetSetting("LAB-i", "OBISettings", "FLD_7", "''")).Value & ""
        End If
    End If

    '--------------------------------------
     Set rstTemp = Nothing
End If
If Len(Trim(lblProfileName.Caption)) > 0 Then DisplayResultForm

errFound:
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation, "MODULE: SETVALUES"
    End If


End Sub

Private Function getNVclass(mSex As String) As String
Dim rstTemp As ADODB.Recordset
Set rstTemp = myConnection.Execute("Select classID from tblabclassifications where instrumentid = '" & InstrumentID & "' and classname = '" & mSex & "'")
If rstTemp.RecordCount > 0 Then
    rstTemp.MoveFirst
    getNVclass = rstTemp.Fields("classid").Value
End If
Set rstTemp = Nothing
End Function

Private Sub setGrid()
Dim mShowSI As String
Dim ctr As Integer
mShowSI = Val(GetSetting("HostForms_SA", "Settings", "ShowSI", "1"))
With msfgResult
    Select Case SectionCode
        Case "CHEM", "PLMO", "PULMO"
            If mShowSI = "1" Then
                .ColWidth(0) = .Width * 0.03 - (3 * 3) 'Fixed
                .ColWidth(1) = .Width * 0.15 - (15 * 3) 'ExamName
                .ColWidth(2) = .Width * 0.15 - (15 * 3) 'Original Result
                .ColWidth(3) = .Width * 0.15 - (15 * 3) 'Edited Result
                .ColWidth(4) = .Width * 0.15 - (15 * 3)  'Normal Values
                .ColWidth(5) = .Width * 0.07 - (7 * 3)  'Status
                .ColWidth(15) = .Width * 0.14 - (14 * 3)
                .ColWidth(16) = .Width * 0.14 - (14 * 3)
                .ColWidth(17) = 0   'DilutionSI
                .ColWidth(18) = 0   'Additional Profile
                .ColWidth(19) = 0 'Formula
                .ColWidth(20) = 0 'High - Low Flags
                .ColWidth(21) = 0   'Exam Notes
                .ColWidth(22) = 0   'Exam Type
                .ColWidth(23) = 0   'Exam Type Values
            Else
                .ColWidth(0) = .Width * 0.05 - (30 * 3)
                .ColWidth(1) = .Width * 0.25 - (30 * 3)
                .ColWidth(2) = .Width * 0.15 - (20 * 3)
                .ColWidth(3) = .Width * 0.15 - (10 * 3)
                .ColWidth(4) = .Width * 0.2 - (10 * 3)
                .ColWidth(5) = .Width * 0.2 - (10 * 3)
                .ColWidth(15) = 0
                .ColWidth(16) = 0
                .ColWidth(17) = 0   'DilutionSI
                .ColWidth(18) = 0   'Additional Profile
                .ColWidth(19) = 0 'Formula
                .ColWidth(20) = 0 'High - Low Flags
                .ColWidth(21) = 0   'Exam Notes
                .ColWidth(22) = 0   'Exam Type
                .ColWidth(23) = 0   'Exam Type Values
            End If
            .ColWidth(6) = 0 'Instrument Exam Name
            .ColWidth(7) = 0 'Dilution
            .ColWidth(8) = 0 'Data Format
            .ColWidth(9) = 0 'Exam ID
            .ColWidth(10) = 0
            .ColWidth(11) = 0
            .ColWidth(12) = 0
            .ColWidth(13) = 0   'FactorOP
            .ColWidth(14) = 0   'UnitSI
            .ColWidth(17) = 0   'DilutionSI
            .ColWidth(18) = 0   'Additional Profile
            .ColWidth(19) = 0 'Formula
            .ColWidth(20) = 0 'High - Low Flags
            .ColWidth(21) = 0   'Exam Notes
            .ColWidth(22) = 0   'Exam Type
        Case "HEMA"
            .ColWidth(0) = .Width * 0.05 - (30 * 3)
            .ColWidth(1) = .Width * 0.25 - (30 * 3)
            .ColWidth(2) = .Width * 0.15 - (20 * 3)
            .ColWidth(3) = .Width * 0.15 - (10 * 3)
            .ColWidth(4) = .Width * 0.2 - (10 * 3)
            .ColWidth(5) = .Width * 0.2 - (10 * 3)
            .ColWidth(6) = 0
            .ColWidth(7) = 0
            .ColWidth(8) = 0
            .ColWidth(9) = 0
            .ColWidth(10) = 0
            .ColWidth(11) = 0
            .ColWidth(12) = 0
            .ColWidth(13) = 0 'FactorOP
            .ColWidth(14) = 0 'UnitSI
            .ColWidth(15) = 0 'ResultSI
            .ColWidth(16) = 0 'NVSI
            .ColWidth(17) = 0 'DilutionSI
            .ColWidth(18) = 0 'Additional Profile
            .ColWidth(19) = 0 'Formula
            .ColWidth(20) = 0 'High - Low Flags
            .ColWidth(21) = 0   'Exam Notes
            .ColWidth(22) = 0   'Exam Type
            .ColWidth(23) = 0   'Exam Type Values
        Case "IMMU"
            .ColWidth(0) = .Width * 0.05 - (30 * 3)
            .ColWidth(1) = .Width * 0.25 - (30 * 3)
            .ColWidth(2) = .Width * 0.15 - (20 * 3)
            .ColWidth(3) = .Width * 0.15 - (10 * 3)
            .ColWidth(4) = .Width * 0.2 - (10 * 3)
            .ColWidth(5) = .Width * 0.2 - (10 * 3)
            .ColWidth(6) = 0
            .ColWidth(7) = 0
            .ColWidth(8) = 0
            .ColWidth(9) = 0
            .ColWidth(10) = 0
            .ColWidth(11) = 0
            .ColWidth(12) = 0
            .ColWidth(13) = 0 'FactorOP
            .ColWidth(14) = 0 'UnitSI
            .ColWidth(15) = 0 'ResultSI
            .ColWidth(16) = 0 'NVSI
            .ColWidth(17) = 0 'DilutionSI
            .ColWidth(18) = 0 'Additional Profile
            .ColWidth(19) = 0 'Formula
            .ColWidth(20) = 0 'High - Low Flags
            .ColWidth(21) = 0   'Exam Notes
            .ColWidth(22) = 0   'Exam Notes
            .ColWidth(23) = 0   'Exam Type Values
    End Select
'    If Val(GetSetting("HostForms_SA", "Settings", "Remarks", "1")) = 1 Then
'        .Height = (.RowHeightMin * 14) + 38
'    Else
'        .Height = (.RowHeightMin * 16) + 38
'    End If
    lblTextDiff.Top = .Top + .Height + 20
    lblDiff.Top = .Top + .Height + 20
    
    For ctr = 1 To .Rows - 1
        .RowHeight(ctr) = 315
    Next ctr
End With
End Sub

Private Sub Form_Resize()
If Val(GetSetting("HostForms_SA", "Settings", "AutoResize", "0")) = 1 Then
    fraTitle.Width = Me.ScaleWidth
    fraBody(0).Width = Me.ScaleWidth
    fraBody(1).Width = Me.ScaleWidth
    fraBody(2).Width = Me.ScaleWidth
    fraBody(3).Width = Me.ScaleWidth
    fraBody(3).Top = Me.ScaleHeight - fraBody(3).Height
    fraBody(1).Height = Me.ScaleHeight - fraBody(0).Height - fraBody(3).Height
    fraBody(2).Height = Me.ScaleHeight - fraBody(0).Height - fraBody(3).Height
    msfgResult.Width = fraBody(2).Width - 200
    cmdClose.Left = fraBody(3).Width - (cmdClose.Width + 50)
    cmdSave.Left = cmdClose.Left - (cmdSave.Width + 50)
    cmdSavePrint.Left = cmdSave.Left - (cmdSavePrint.Width + 50)
    cmdSavePreview.Left = cmdSavePrint.Left - (cmdSavePreview.Width + 50)
    lblTextDiff.Width = fraBody(2).Width * 0.15
    lblDiff.Width = fraBody(2).Width * 0.05
    lblDiff.Left = lblTextDiff.Left + lblTextDiff.Width
    
    If lblRemarks2.Visible Then
        lblRemarks.Width = fraBody(2).Width * 0.35
        txtRemarks.Width = fraBody(2).Width * 0.35
        lblRemarks.Left = lblDiff.Left + lblDiff.Width + 300
        txtRemarks.Left = lblRemarks.Left
        lblRemarks.Top = fraBody(2).Height - (lblRemarks.Height + txtRemarks.Height) - 600
        txtRemarks.Top = fraBody(2).Height - (txtRemarks.Height) - 600
    Else
        lblRemarks.Width = fraBody(2).Width * 0.77
        txtRemarks.Width = fraBody(2).Width * 0.77
        lblRemarks.Left = lblDiff.Left + lblDiff.Width + 300
        txtRemarks.Left = lblRemarks.Left
        lblRemarks.Top = fraBody(2).Height - (lblRemarks.Height + txtRemarks.Height) - 600
        txtRemarks.Top = fraBody(2).Height - (txtRemarks.Height) - 600
    End If
    
    lblRemarks2.Width = fraBody(2).Width * 0.35
    txtRemarks2.Width = fraBody(2).Width * 0.35
    lblRemarks2.Left = lblRemarks.Left + lblRemarks.Width + 600
    txtRemarks2.Left = lblRemarks2.Left
    lblRemarks2.Top = lblRemarks.Top
    txtRemarks2.Top = txtRemarks.Top
   
    
    lblTextDiff.Top = lblRemarks.Top
    lblDiff.Top = lblRemarks.Top
    msfgResult.Height = lblRemarks.Top - 300
    Label16.Width = fraBody(1).Width
    Shape1.Left = (fraBody(2).Width / 2) - (Shape1.Width / 2)
    setPosDemo
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rstPatient = Nothing
    Set rstPatient2 = Nothing
    Set rstSavedResult = Nothing
    Set rstClass = Nothing
'    Call WheelUnHook(Me.hWnd)
End Sub

Private Sub Label2_Click()
        yClassForms.showSearchCode 5, Mid(txtMedtech.Text, 1, Len(txtMedtech.Text) - txtMedtech.SelLength), , , , True
        txtMedtech.Text = yClassForms.LastSearchName
End Sub

Private Sub Label1_Click()
Dim tempStr As String
        tempStr = txtPathologist.Text
        yClassForms.showSearchCode 7, ""
        If Len(Trim(yClassForms.LastSearchName)) > 0 Then
            txtPathologist.Text = yClassForms.LastSearchName
            myConnection.Execute "update tbLabPathologists set [Default] = False"
            myConnection.Execute "update tbLabPathologists set [Default] = True where Pathologist = '" & txtPathologist.Text & "'"
        Else
            txtPathologist.Text = tempStr
        End If
End Sub

Private Sub Label3_Click()
        yClassForms.showSearchCode 6, Mid(txtPhysician.Text, 1, Len(txtPhysician.Text) - txtPhysician.SelLength), , , , True
        txtPhysician.Text = yClassForms.LastSearchName
End Sub

Private Sub lblClassB_Click()
    txtClassID.Text = ""
    yClassForms.showSearchCode 8, txtClassID
    txtClassID.Text = yClassForms.LastsearchID
    lblClassName.Caption = yClassForms.LastSearchName
End Sub

Private Sub lblClassName_Change()
If Len(Trim(lblClassName.Caption)) > 0 Then
    If DBServerType = "LOCAL DATABASE" Then
        Set rstClass = myConnection.Execute("Select * from tblabClassExams where Active = True and InstrumentID = '" & InstrumentID & "' and ClassID ='" & txtClassID.Text & "'")
    Else
        Set rstClass = myConnection.Execute("Select * from tblabClassExams where Active = 1 and InstrumentID = '" & InstrumentID & "' and ClassID ='" & txtClassID.Text & "'")
    End If
    setNValues
Else
    setNValues True
End If
If Len(Trim(lblProfileName.Caption)) < 1 Or Len(Trim(lblClassName.Caption)) < 1 Or getMandatoryFields = False Then
    cmdSavePreview.Enabled = False
    cmdSavePrint.Enabled = False
    cmdSave.Enabled = False
Else
    cmdSavePreview.Enabled = True
    cmdSavePrint.Enabled = True
    cmdSave.Enabled = True
End If
setComputedValue
End Sub

Private Sub lblDiff_Change()
If validateDiff Then
    lblTextDiff.BackStyle = 1
    lblTextDiff.BorderStyle = 1
    lblTextDiff.BackColor = &H80000005
    lblTextDiff.ForeColor = &H0&
Else
    lblTextDiff.BackStyle = 1
    lblTextDiff.BorderStyle = 1
    lblTextDiff.BackColor = &HC0&
    lblTextDiff.ForeColor = &HFFFFFF
End If
End Sub

Private Sub lblProfileidB_Click()
    txtProfileID.Text = ""
    yClassForms.showSearchCode 1, txtProfileID
    txtProfileID.Text = yClassForms.LastsearchID
    lblProfileName.Caption = yClassForms.LastSearchName
    If Len(Trim(lblProfileName.Caption)) > 0 Then
        DisplayResultForm
    End If
End Sub

Private Sub lblProfileName_Change()
lblProfileName.ToolTipText = lblProfileName.Caption
If Len(Trim(lblProfileName.Caption)) < 1 Or Len(Trim(lblClassName.Caption)) < 1 Or getMandatoryFields = False Then
    cmdSavePreview.Enabled = False
    cmdSavePrint.Enabled = False
    cmdSave.Enabled = False
Else
    cmdSavePreview.Enabled = True
    cmdSavePrint.Enabled = True
    cmdSave.Enabled = True
    lblProfileName.ToolTipText = lblProfileName.Caption
End If
End Sub

Private Sub cmdSavePrint_Click()
Dim myArray() As String
Dim ctr As Integer
Dim tempStr As String
Dim tempTag As String
ctr = 0
tempStr = ""
If Len(Trim(lblClassName.Caption)) > 0 And Len(Trim(lblProfileName.Caption)) > 0 And fraBody(1).Visible = True And Len(Trim(txtLastName.Text)) > 0 Then
    If NewResult Then
        If Not HasRefNum Then saveNewResult
    End If
    SaveResult True
    If validateDiff Or Val(GetSetting("HostForms_SA", "Settings", "ByPassDiff", "0")) = 1 Then
        Me.MousePointer = 11
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        If ShowPrintCopy Then
            tempStr = yClassForms.showPrintCopies
            myArray = Split(tempStr, "|")
            If Val(myArray(0)) > 0 Then
                For ctr = 1 To Val(myArray(0))
                    If OpenMainReport(App.Path & "\reports\" & getDefaultTemplate(txtRefNum.Text), txtRefNum.Text & "", CStr(ctr)) <> "Error" Then
                        ShowReportViewer False, "Result Print Preview", True, txtRefNum.Text, 1, CStr(ctr)
                    End If
                Next ctr
            End If
        Else
            If OpenMainReport(App.Path & "\reports\" & getDefaultTemplate(txtRefNum.Text), txtRefNum.Text & "", "") <> "Error" Then
                ShowReportViewer False, "Result Print Preview", True, txtRefNum.Text
            End If
        End If
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        If ShowPrintCopy Then
            If myArray(1) = 1 Then
                tempTag = myArray(3)
                For ctr = 1 To Val(myArray(2))
                    If OpenMainReport(App.Path & "\reports\" & getDefaultTemplate(txtRefNum.Text), txtRefNum.Text & "", "T" & String(3 - Len(CStr(ctr)), "0") & CStr(ctr) & "-" & tempTag) <> "Error" Then
                        If myArray(1) = 1 Then
                            ShowReportViewer False, "Result Print Preview", True, txtRefNum.Text, 1, "T" & String(3 - Len(CStr(ctr)), "0") & CStr(ctr) & "-" & tempTag
                        End If
                    End If
                Next ctr
            End If
        End If
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Me.MousePointer = 0
        If (yMultiCurrent - yMultiStart) + 1 = (yMultiTotal - yMultiStart) + 1 Then
            Me.MousePointer = 0
            ReturnValueInt = yMultiTotal + 1
            If tempStr <> "0|0|0|0" Then Unload Me
        Else
            If cmdNext.Visible Then
                If tempStr <> "0|0|0|0" Then
                    cmdNext_Click
                End If
            End If
        End If
    End If
Else
    MsgBox "Cannot Save Result...", vbOKOnly + vbInformation, "Result Entry"
End If
End Sub



Private Sub lblRemarks_Click()
' If lblRemarks.Visible = True Then
'    lblRemarks.Width = "8415"
' End If
End Sub

Private Sub lblTitle_Change()
    Me.Caption = lblTitle.Caption
End Sub

Private Sub hideBlankInstrumentResults()
Dim ctr As Integer
With msfgResult
'    .RowHeightMin = 0
    For ctr = 1 To .Rows - 1
        If Len(Trim(.TextMatrix(ctr, 3))) = 0 Then
            If .RowHeight(ctr) = 0 Then
                .RowHeight(ctr) = 315
                .Row = 1
            Else
                .RowHeight(ctr) = 0
            End If
        End If
    Next ctr
End With
End Sub

Private Sub msfgResult_Click()
With msfgResult
    If .MouseRow = 0 And .MouseCol = 0 Then
        hideBlankInstrumentResults
    End If
    .Visible = False
    If .TextMatrix(.Row, 22) = "NV" Or Len(Trim(.TextMatrix(.Row, 22))) = 0 Then
        .Col = 3
        If .RowHeight(.Row) <= 0 Then txtEntry.Visible = False Else txtEntry.Visible = True
        cmdStringValue.Visible = False
        cboResult.Visible = False
        If txtEntry.Visible Then
            txtEntry.SetFocus
            txtEntry.Text = .TextMatrix(.Row, 3)
            txtEntry.Width = .CellWidth
            txtEntry.Height = .CellHeight
            txtEntry.Top = .CellTop + .Top
            txtEntry.Left = .CellLeft + .Left
        End If
        If Len(Trim(lblClassName.Caption)) > 0 Then setNValuesItem .Row
        If Len(Trim(.TextMatrix(.Row, 19))) = 0 Then
            txtEntry.Locked = False
            txtEntry.BackColor = &HC0E0FF
        Else
            txtEntry.Locked = True
            txtEntry.BackColor = &H80000005
        End If
    ElseIf .TextMatrix(.Row, 22) = "SV" Then
        .Col = 3
        txtEntry.Visible = False
        cmdStringValue.Visible = True
        cboResult.Visible = False
        cmdStringValue.SetFocus
        cmdStringValue.Width = .CellWidth
        cmdStringValue.Height = .CellHeight
        cmdStringValue.Top = .CellTop + .Top
        cmdStringValue.Left = .CellLeft + .Left
    ElseIf .TextMatrix(.Row, 22) = "MV" Then
        .Col = 3
        txtEntry.Visible = False
        cmdStringValue.Visible = False
        cboResult.Visible = True
        cboResult.Width = .CellWidth
        cboResult.Top = .CellTop + .Top
        cboResult.Left = .CellLeft + .Left
        FillCBOResult .TextMatrix(.Row, 23)
        cboResult.ListIndex = findCboIndex(cboResult, .TextMatrix(.Row, .Col))
        cboResult.SetFocus
    End If
    .Visible = True
End With
End Sub

Private Sub FillCBOResult(myStr As String)
Dim rstSelection As ADODB.Recordset
Set rstSelection = myConnection.Execute("Select SelValues from tbSelectionValues where SelType = '" & myStr & "' order by SelSortOrder")
If rstSelection.RecordCount > 0 Then
    rstSelection.MoveFirst
    cboResult.Clear
    cboResult.AddItem ""
    Do While Not rstSelection.EOF
        cboResult.AddItem rstSelection.Fields("SelValues").Value
        rstSelection.MoveNext
    Loop
End If
End Sub

Private Sub msfgResult_EnterCell()
    msfgResult_Click
End Sub

Private Sub msfgResult_GotFocus()
    msfgResult_Click
End Sub

Private Sub msfgResult_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If msfgResult.MouseCol = 0 And msfgResult.MouseRow = 0 Then
    msfgResult.ToolTipText = "Click here to hide/show blank rows."
ElseIf msfgResult.MouseCol = 3 Then
    msfgResult.ToolTipText = msfgResult.TextMatrix(msfgResult.MouseRow, 0)
Else
    msfgResult.ToolTipText = ""
End If
End Sub

Private Sub msfgResult_Scroll()
With msfgResult
        If .RowIsVisible(.Row) Then
            If Val(GetSetting("HostForms_SA", "Settings", "Remarks", "1")) = 1 Then
                If .Row - .TopRow = 13 Then
                    txtEntry.Visible = False
                    cmdStringValue.Visible = False
                    cboResult.Visible = False
                Else
                    msfgResult_Click
                End If
            Else
                If .Row - .TopRow = 15 Then
                    txtEntry.Visible = False
                    cmdStringValue.Visible = False
                    cboResult.Visible = False
                Else
                    msfgResult_Click
                End If
            End If
        Else
            txtEntry.Visible = False
            cmdStringValue.Visible = False
            cboResult.Visible = False
        End If
'lblRemarks.Caption = .Row & " / " & .TopRow
End With
End Sub

Private Sub txtAge_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtAge_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtAge_LostFocus()
    txtAge.BackColor = &H80000005
End Sub

Private Sub txtBirthDate_Change()
Dim xIsDateMonths As String
Dim xYears As String
Dim xDays As String
Dim xMonths As String
Dim tYears As Integer
Dim tDays As Integer
Dim tMonths As Integer
Dim tempAge As Integer
If Val(GetSetting("HostForms_SA", "Settings", "ComputeAge", "1")) = 1 Then
    If IsDate(txtBirthDate) Then
        xYears = DateDiff("yyyy", txtBirthDate.Text, Date)
        xMonths = DateDiff("m", txtBirthDate.Text, Date)
        xDays = DateDiff("d", txtBirthDate.Text, Date)
        If xDays = 0 And xMonths = 0 And xYears = 0 Then
            txtAge.Text = "NEWBORN"
        Else
            If xYears = 1 Then
                If Month(Now) = Month(txtBirthDate.Text) And Day(Now) >= Day(txtBirthDate.Text) Then
                    txtAge.Text = xYears
                ElseIf Month(Now) = Month(txtBirthDate.Text) And Day(Now) < Day(txtBirthDate.Text) Then
                    txtAge.Text = xMonths - 1 & IIf(xMonths = 1, " MONTH", " MONTHS")
                ElseIf Month(Now) < Month(txtBirthDate.Text) And Day(Now) >= Day(txtBirthDate.Text) Then
                    txtAge.Text = xMonths & IIf(xMonths = 1, " MONTH", " MONTHS")
                ElseIf Month(Now) < Month(txtBirthDate.Text) And Day(Now) < Day(txtBirthDate.Text) Then
                    txtAge.Text = (xMonths - 1) & IIf(xMonths = 1, " MONTH", " MONTHS")
                ElseIf Month(Now) > Month(txtBirthDate.Text) And Day(Now) >= Day(txtBirthDate.Text) Then
                    txtAge.Text = xYears
                ElseIf Month(Now) > Month(txtBirthDate.Text) And Day(Now) < Day(txtBirthDate.Text) Then
                    txtAge.Text = xYears
                End If
            ElseIf xYears > 1 Then
                If Month(Now) = Month(txtBirthDate.Text) And Day(Now) >= Day(txtBirthDate.Text) Then
                    txtAge.Text = xYears
                ElseIf Month(Now) = Month(txtBirthDate.Text) And Day(Now) < Day(txtBirthDate.Text) Then
                    txtAge.Text = xYears - 1
                ElseIf Month(Now) < Month(txtBirthDate.Text) And Day(Now) >= Day(txtBirthDate.Text) Then
                    txtAge.Text = xYears - 1
                ElseIf Month(Now) < Month(txtBirthDate.Text) And Day(Now) < Day(txtBirthDate.Text) Then
                    txtAge.Text = xYears - 1
                ElseIf Month(Now) > Month(txtBirthDate.Text) And Day(Now) >= Day(txtBirthDate.Text) Then
                    txtAge.Text = xYears
                ElseIf Month(Now) > Month(txtBirthDate.Text) And Day(Now) < Day(txtBirthDate.Text) Then
                    txtAge.Text = xYears
                End If
            ElseIf xYears < 1 Then
                If Month(Now) = Month(txtBirthDate.Text) And Day(Now) >= Day(txtBirthDate.Text) Then
                    If xDays < 1 Then
                        txtAge.Text = ""
                    Else
                        txtAge.Text = xDays & IIf(xDays = 1, " DAY", " DAYS")
                    End If
                ElseIf Month(Now) = Month(txtBirthDate.Text) And Day(Now) < Day(txtBirthDate.Text) Then
                    txtAge.Text = ""
                ElseIf Month(Now) < Month(txtBirthDate.Text) And Day(Now) >= Day(txtBirthDate.Text) Then
                    If xDays < 1 Then
                        txtAge.Text = ""
                    Else
                        txtAge.Text = xDays & IIf(xDays = 1, " DAY", " DAYS")
                    End If
                ElseIf Month(Now) < Month(txtBirthDate.Text) And Day(Now) < Day(txtBirthDate.Text) Then
                    If xDays < 1 Then
                        txtAge.Text = ""
                    Else
                        txtAge.Text = xDays & IIf(xDays = 1, " DAY", " DAYS")
                    End If
                ElseIf Month(Now) > Month(txtBirthDate.Text) And Day(Now) >= Day(txtBirthDate.Text) Then
                    If xMonths < 1 Then
                        txtAge.Text = ""
                    Else
                        txtAge.Text = xMonths & IIf(xMonths = 1, " MONTH", " MONTHS")
                    End If
                ElseIf Month(Now) > Month(txtBirthDate.Text) And Day(Now) < Day(txtBirthDate.Text) Then
                    If xDays < 1 Then
                        txtAge.Text = ""
                    Else
                        txtAge.Text = xMonths & IIf(xMonths = 1, " MONTH", " MONTHS")
                    End If
                End If
            End If
        End If
'        If xDays < 30 And xDays > 0 Then
'            txtAge.Text = xDays & IIf(xDays = 1, " DAY", " DAYS")
'        ElseIf xMonths < 12 And xMonths > 0 Then
'            If Val(Format(Now, "dd")) < Val(Format(txtBirthDate.Text, "dd")) Then
'                txtAge.Text = xMonths - 1 & IIf(xMonths - 1 = 1, " MONTH", " MONTHS")
'            Else
'                txtAge.Text = xMonths & IIf(xMonths = 1, " MONTH", " MONTHS")
'            End If
'        ElseIf xMonths > 12 Then
'            txtAge.Text = Int(xMonths / 12)
'            If Month(Now) = Month(txtBirthDate.Text) Then
'                If Val(Format(Now, "dd")) < Val(Format(txtBirthDate.Text, "dd")) Then
'                    txtAge.Text = Int(xMonths / 12) - 1
'                End If
'            End If
'        ElseIf xDays = 0 And xMonths = 0 And xYears = 0 Then
'            txtAge.Text = "NEWBORN"
'        ElseIf xMonths = 12 Then
'            If Val(Format(Now, "dd")) >= Val(Format(txtBirthDate.Text, "dd")) Then
'                txtAge.Text = "1"
'            Else
'                txtAge.Text = xMonths - 1 & " MONTHS"
'            End If
'        Else
'            txtAge.Text = ""
'        End If
    Else
        txtAge.Text = ""
    End If
End If
End Sub

Private Sub txtBirthDate_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtBirthDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtBirthDate_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtBirthDate_LostFocus()
    txtBirthDate.BackColor = &H80000005
End Sub

Private Sub txtCivilStatus_Change()
If Len(Trim(lblProfileName.Caption)) < 1 Or Len(Trim(lblClassName.Caption)) < 1 Or getMandatoryFields = False Then
    cmdSavePreview.Enabled = False
    cmdSavePrint.Enabled = False
    cmdSave.Enabled = False
Else
    cmdSavePreview.Enabled = True
    cmdSavePrint.Enabled = True
    cmdSave.Enabled = True
End If
End Sub

Private Sub txtCivilStatus_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtCivilStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        keybd_event VK_TAB, 0, 0, 0
    ElseIf KeyCode = vbKeyF2 Then
        If GetSetting("HostForms_SA", "SelList", "cCivil Status", "0") = 1 Then
            yClassForms.showSearchCode 11, GetSetting("HostForms_SA", "SelList", "Civil Status", "Civil Status"), , lblCivilStatus.Caption, , True
            txtCivilStatus.Text = yClassForms.LastSearchName
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If txtCivilStatus.Locked Then
            txtCivilStatus.Text = ""
        End If
    End If
End Sub

Private Sub txtCivilStatus_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtCivilStatus_LostFocus()
    txtCivilStatus.BackColor = &H80000005
End Sub

Private Sub txtClassID_Change()
    lblClassName.Caption = ""
End Sub

Private Sub txtClassID_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtClassID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    yClassForms.showSearchCode 8, txtClassID
    txtClassID.Text = yClassForms.LastsearchID
    lblClassName.Caption = yClassForms.LastSearchName
    If Len(Trim(lblClassName.Caption)) > 0 Then
    End If
End If
End Sub

Private Sub txtClassID_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtClassID_LostFocus()
    txtClassID.BackColor = &H80000005
End Sub

Private Sub txtEntry_Change()
With msfgResult
    .TextMatrix(.Row, .Col) = txtEntry.Text
'    setNValuesItem .Row
    setComputedValue
    setSIvalues
    setNValues
'    setModifiedColor .Row, .TextMatrix(.Row, 2), .TextMatrix(.Row, 3)
    lblDiff.Caption = getTotalDiff
End With
End Sub

Private Sub txtEntry_GotFocus()
    highLight
    highLight_bg
    With msfgResult
        If Len(Trim(.TextMatrix(.Row, 19))) = 0 Then
            txtEntry.Locked = False
            txtEntry.BackColor = &HC0E0FF
        Else
            txtEntry.Locked = True
            txtEntry.BackColor = &H80000005
        End If
    End With
End Sub

Private Sub txtEntry_KeyDown(KeyCode As Integer, Shift As Integer)
With msfgResult
If KeyCode = vbKeyReturn Then
        If .Row = .Rows - 1 Then
            .Row = .Row
'            keybd_event VK_TAB, 0, 0, 0
            If txtRemarks.Visible Then txtRemarks.SetFocus
            txtEntry.Visible = False
            
            'msfgResult_EnterCell
        Else
            .Row = .Row + 1
            msfgResult_EnterCell
        End If
        highLight
End If
If KeyCode = vbKeyRight And (txtEntry.SelStart = Len(txtEntry.Text) Or txtEntry.SelLength = Len(txtEntry.Text)) And .Col <> .Cols - 1 Then
    .SetFocus
    .Col = .Col + 1
    msfgResult_EnterCell
End If
If KeyCode = vbKeyLeft And .Col <> 1 And txtEntry.SelStart = 0 Then
    .SetFocus
    .Col = .Col - 1
    msfgResult_EnterCell
End If
If KeyCode = vbKeyUp And .Row <> 1 Then
    .SetFocus
    .Row = .Row - 1
    msfgResult_EnterCell
End If
If KeyCode = vbKeyDown And .Row <> .Rows - 1 Then
    .Row = .Row + 1
    .SetFocus
    msfgResult_EnterCell
End If
If KeyCode = vbKeyF3 Then
        .TextMatrix(.Row, 0) = yClassForms.showInputBox(.TextMatrix(.Row, 1) & " - String Value", , , .TextMatrix(.Row, 0), 50, , , True)
        If Len(Trim(.TextMatrix(.Row, 0))) > 0 Then
            .Col = 0
            .CellBackColor = .CellForeColor
        Else
            .Col = 0
            .CellBackColor = .BackColorFixed
        End If
        .Visible = False
        .Visible = True
End If
If KeyCode = vbKeyF2 Then
    If IsNumeric(.TextMatrix(.Row, 2)) Then txtEntry.Text = Val(.TextMatrix(.Row, 2) * Val(.TextMatrix(.Row, 7)))
    txtEntry.Text = Format(txtEntry.Text, IIf((Len(Trim(.TextMatrix(.Row, 8))) > 0), .TextMatrix(.Row, 8), ""))
    'If txtEntry.SelLength > 0 Then txtEntry.SelStart = 0
End If
If KeyCode = vbKeyF10 Then
    If Len(Trim(.TextMatrix(.Row, 21))) > 0 Then
        MsgBox .TextMatrix(.Row, 21), vbOKOnly, "Details"
    End If
End If
If KeyCode = vbKeyF4 Then
    hideBlankInstrumentResults
    If .RowHeight(.Row) <= 0 Then
        txtEntry.Visible = False
        If .RowIsVisible(.Row) Then
            .Row = .Rows - 1
        End If
    End If
End If

End With
End Sub

Private Sub setModifiedColor(mRow As Integer, mOrgValue As String, mNewValue As String)
Dim ctr As Integer
Dim orgValue As Double
Dim NewValue As Double
With msfgResult
If IsNumeric(mOrgValue) And IsNumeric(mNewValue) Then
    orgValue = CDbl(mOrgValue)
    NewValue = CDbl(mNewValue)
    If orgValue <> NewValue Then
        .Row = mRow
        For ctr = 1 To 4
            .Col = ctr
            .CellBackColor = &HC0C0FF       '&HC0C0C0
        Next ctr
    Else
        .Row = mRow
        For ctr = 1 To 4
            .Col = ctr
            .CellBackColor = &HFFFFFF
        Next ctr
    End If
ElseIf mOrgValue = "" Or mNewValue = "" Then
    If Val(mOrgValue) <> Val(mNewValue) Then
        .Row = mRow
        For ctr = 1 To 4
            .Col = ctr
            .CellBackColor = &HC0C0FF
        Next ctr
    Else
        .Row = mRow
        For ctr = 1 To 4
            .Col = ctr
            .CellBackColor = &HFFFFFF
        Next ctr
    End If
Else
    .Row = mRow
    For ctr = 1 To 4
        .Col = ctr
        .CellBackColor = &HFFFFFF
    Next ctr
End If
.Col = 3
End With
End Sub

Private Sub txtEntry_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtEntry_LostFocus()
    txtEntry.BackColor = &H80000005
End Sub

Private Sub txtFirstName_Change()
    lblName.Caption = txtLastName.Text & ", " & txtFirstName.Text & " " & txtMiddleName.Text
End Sub

Private Sub txtFirstName_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtFirstName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtFirstName_LostFocus()
    txtFirstName.BackColor = &H80000005
End Sub

Private Sub txtLastName_Change()
    lblName.Caption = txtLastName.Text & ", " & txtFirstName.Text & " " & txtMiddleName.Text
End Sub

Private Sub txtLastName_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtLastName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtLastName_LostFocus()
    txtLastName.BackColor = &H80000005
End Sub

Private Sub txtmedtech_Change()
  Dim Location As Long
  lstMedtech.ListIndex = SendMessage(lstMedtech.hWnd, _
        LB_FINDSTRING, -1, ByVal CStr(txtMedtech.Text))
  If Not freeDelete Then
    If lstMedtech.ListIndex <> -1 Then
      Location = txtMedtech.SelStart
      txtMedtech.Text = lstMedtech.Text
      txtMedtech.SelStart = Location
      txtMedtech.SelLength = Len(txtMedtech) - Location
    End If
  End If
End Sub

Private Sub txtMedtech_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtMedTech_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Temp As String
Select Case KeyCode
    Case 8
'        If txtPhysician.SelLength <> 0 Then
''          txtPhysician.Text = Mid$(txtPhysician.Text, 1, txtPhysician.SelStart - 1)
          KeyCode = 0
'        End If
        freeDelete = True
    Case vbKeyDelete
        If txtMedtech.SelLength <> 0 And _
           txtMedtech.SelStart <> 0 Then
           txtMedtech.SelText = ""
           KeyCode = 0
        End If
        freeDelete = True
    Case vbKeyReturn
        keybd_event VK_TAB, 0, 0, 0
        freeDelete = False
    Case vbKeyF2
        yClassForms.showSearchCode 5, Mid(txtMedtech.Text, 1, Len(txtMedtech.Text) - txtMedtech.SelLength), , , , True
        txtMedtech.Text = yClassForms.LastSearchName
    Case Else
        freeDelete = False
End Select
End Sub

Private Sub txtmedtech_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtMedtech_LostFocus()
    txtMedtech.BackColor = &H80000005
End Sub

Private Sub txtMiddleName_Change()
    lblName.Caption = txtLastName.Text & ", " & txtFirstName.Text & " " & txtMiddleName.Text
End Sub

Private Sub txtMiddleName_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtMiddleName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtMiddleName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtMiddleName_LostFocus()
    txtMiddleName.BackColor = &H80000005
End Sub

Private Sub txtPathologist_GotFocus()
    highLight_bg
End Sub

Private Sub txtPathologist_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tempStr As String
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
If KeyCode = vbKeyF2 Then
        tempStr = txtPathologist.Text
        yClassForms.showSearchCode 7, ""
        If Len(Trim(yClassForms.LastSearchName)) > 0 Then
            txtPathologist.Text = yClassForms.LastSearchName
            If DBServerType = "MSSQL SERVER" Then
                myConnection.Execute "update tbLabPathologists set [Default] = 0"
                myConnection.Execute "update tbLabPathologists set [Default] = 1 where Pathologist = '" & txtPathologist.Text & "'"
            Else
                myConnection.Execute "update tbLabPathologists set [Default] = False"
                myConnection.Execute "update tbLabPathologists set [Default] = True where Pathologist = '" & txtPathologist.Text & "'"
            End If
        Else
            txtPathologist.Text = tempStr
        End If
End If
End Sub

Private Sub txtPathologist_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtPathologist_LostFocus()
    txtPathologist.BackColor = &H80000005
End Sub

Private Sub txtPatientID_Change()
If Len(Trim(lblProfileName.Caption)) < 1 Or Len(Trim(lblClassName.Caption)) < 1 Or getMandatoryFields = False Then
    cmdSavePreview.Enabled = False
    cmdSavePrint.Enabled = False
    cmdSave.Enabled = False
Else
    cmdSavePreview.Enabled = True
    cmdSavePrint.Enabled = True
    cmdSave.Enabled = True
End If
If Trim(Right(txtRefNum.Text, 1)) = "N" Then
    If Not HasRefNum Then clearValuesN
End If
End Sub

Private Sub txtPatientID_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtPatientID_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tempStr As String
If KeyCode = vbKeyReturn Then
    If EnableSearchID = 1 And Not HasRefNum Then
        tempStr = yClassForms.showPIDsearch(txtPatientID.Text, True)
        If Len(Trim((tempStr))) > 0 Then
            AssignPIDValues tempStr, txtLastName.Text, txtFirstName.Text
            'keybd_event VK_TAB, 0, 0, 0
        Else
            keybd_event VK_TAB, 0, 0, 0
        End If
    Else
        keybd_event VK_TAB, 0, 0, 0
    End If
ElseIf KeyCode = vbKeyF2 Then
'    setValues
    If txtPatientID.Visible Then txtPatientID.SetFocus
End If
End Sub

Private Sub AssignPIDValues(myPIDSearchValue As String, OrgLastName As String, OrgFirstName As String)
Dim myArray() As String
myArray = Split(myPIDSearchValue, "|")
If Len(Trim(myPIDSearchValue)) > 0 Then
    If myArray(1) <> OrgLastName And myArray(2) <> OrgFirstName Then
        If Val(GetSetting("LAB-i", "OBISettings", "UseExDatabase", "0")) = 1 Then
            txtLastName.Text = myArray(1)
            txtFirstName.Text = myArray(2)
            txtMiddleName.Text = myArray(3)
            txtBirthDate.Text = myArray(4)
            txtCivilStatus.Text = myArray(5)
            txtAge.Text = myArray(6)
            txtSex.Text = myArray(7)
            txtSpecimen.Text = myArray(8)
            dtpResultDate.Value = Format(myArray(12), "mm/dd/yyyy")
            dtpResultTime.Value = Format(myArray(12), "HH:nn:ss")
            dtpRequestDate.Value = Format(myArray(13), "mm/dd/yyyy")
            dtpRequestTime.Value = Format(myArray(13), "HH:nn:ss")
            txtPhysician.Text = myArray(9)
            txtWard.Text = myArray(10)
            txtRoom.Text = myArray(11)
            If txtSpecimenID.Visible Then txtSpecimenID.SetFocus
        Else
            If MsgBox("Are sure you want to use this Name?" & vbCrLf & _
                "Original Name: " & OrgLastName & ", " & OrgFirstName & vbCrLf & _
                "New Name: " & myArray(1) & ", " & myArray(2), vbYesNo + vbQuestion + vbDefaultButton2, "Use Patient ID") = vbYes Then
                txtLastName.Text = myArray(1)
                txtFirstName.Text = myArray(2)
                txtMiddleName.Text = myArray(3)
                txtBirthDate.Text = myArray(4)
                txtCivilStatus.Text = myArray(5)
                txtAge.Text = myArray(6)
                txtSex.Text = myArray(7)
                txtSpecimen.Text = myArray(8)
                dtpResultDate.Value = Format(myArray(12), "mm/dd/yyyy")
                dtpResultTime.Value = Format(myArray(12), "HH:nn:ss")
                dtpRequestDate.Value = Format(myArray(13), "mm/dd/yyyy")
                dtpRequestTime.Value = Format(myArray(13), "HH:nn:ss")
                txtPhysician.Text = myArray(9)
                txtWard.Text = myArray(10)
                txtRoom.Text = myArray(11)
                If txtSpecimenID.Visible Then txtSpecimenID.SetFocus
            End If
        End If
    Else
        txtLastName.Text = myArray(1)
        txtFirstName.Text = myArray(2)
        txtMiddleName.Text = myArray(3)
        txtBirthDate.Text = myArray(4)
        txtCivilStatus.Text = myArray(5)
        txtAge.Text = myArray(6)
        txtSex.Text = myArray(7)
        txtSpecimen.Text = myArray(8)
        dtpResultDate.Value = Format(myArray(12), "mm/dd/yyyy")
        dtpResultTime.Value = Format(myArray(12), "HH:nn:ss")
        dtpRequestDate.Value = Format(myArray(13), "mm/dd/yyyy")
        dtpRequestTime.Value = Format(myArray(13), "HH:nn:ss")
        txtPhysician.Text = myArray(9)
        txtWard.Text = myArray(10)
        txtRoom.Text = myArray(11)
        If txtSpecimenID.Visible Then txtSpecimenID.SetFocus
    End If
End If
End Sub

Private Sub txtPatientID_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtPatientID_LostFocus()
   txtPatientID.BackColor = &H80000005
End Sub

Private Sub txtPhysician_Change()
  Dim Location As Long
  lstPhysician.ListIndex = SendMessage(lstPhysician.hWnd, _
        LB_FINDSTRING, -1, ByVal CStr(txtPhysician.Text))
  If Not freeDelete Then
    If lstPhysician.ListIndex <> -1 Then
      Location = txtPhysician.SelStart
      txtPhysician.Text = lstPhysician.Text
      txtPhysician.SelStart = Location
      txtPhysician.SelLength = Len(txtPhysician) - Location
    End If
  End If
End Sub

Private Sub txtPhysician_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtPhysician_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Temp As String
Select Case KeyCode
    Case 8
'        If txtPhysician.SelLength <> 0 Then
''          txtPhysician.Text = Mid$(txtPhysician.Text, 1, txtPhysician.SelStart - 1)
          KeyCode = 0
'        End If
        freeDelete = True
    Case vbKeyDelete
        If txtPhysician.SelLength <> 0 And _
           txtPhysician.SelStart <> 0 Then
           txtPhysician.SelText = ""
           KeyCode = 0
        End If
        freeDelete = True
    Case vbKeyReturn
        keybd_event VK_TAB, 0, 0, 0
        freeDelete = False
    Case vbKeyF2
        yClassForms.showSearchCode 6, Mid(txtPhysician.Text, 1, Len(txtPhysician.Text) - txtPhysician.SelLength), , , , True
        txtPhysician.Text = yClassForms.LastSearchName
    Case Else
        freeDelete = False
End Select
End Sub

Private Sub txtPhysician_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtPhysician_LostFocus()
    txtPhysician.BackColor = &H80000005
End Sub

Private Sub txtProfileID_Change()
    lblProfileName.Caption = ""
    DisplayResultForm False
End Sub

Private Sub txtProfileID_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtProfileID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    yClassForms.showSearchCode 1, txtProfileID
    txtProfileID.Text = yClassForms.LastsearchID
    lblProfileName.Caption = yClassForms.LastSearchName
    If Len(Trim(lblProfileName.Caption)) > 0 Then
        DisplayResultForm
    End If
'ElseIf KeyCode = vbKeyF2 Then
'    If txtProfileID = "0" Then
'        AdditionalProfiles = yClassForms.showAddProfile(AdditionalProfiles)
'        If Len(Trim(lblProfileName.Caption)) > 0 Then
'            DisplayResultForm
'        End If
'    Else
'        AdditionalProfiles = ""
'    End If
End If
'If txtProfileID = "0" Then
'    If Len(Trim(AdditionalProfiles)) > 0 Then
'        lblProfileName.Caption = "Others(" & AdditionalProfiles & ")"
'    Else
'        lblProfileName.Caption = "Others"
'    End If
'End If
'If txtProfileID.Text <> "0" Then AdditionalProfiles = ""
'setAdditionalProfiles AdditionalProfiles
End Sub

Private Sub txtProfileID_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtProfileID_LostFocus()
    txtProfileID.BackColor = &H80000005
End Sub


Private Sub txtRemarks_GotFocus()
    txtRemarks.BackColor = &HC0E0FF
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tempStr As String
If KeyCode = vbKeyF2 Then
    If txtRemarks.SelStart > 0 Then
        txtRemarks.SelStart = txtRemarks.SelStart - 1
        txtRemarks.SelLength = 1
        tempStr = txtRemarks.SelText
    Else
        tempStr = ""
    End If
    If tempStr = " " Then tempStr = "%"
    yClassForms.showSearchCode 9, tempStr, False, , , True
    If Len(yClassForms.LastSearchName) > 0 Then
        If tempStr = "%" Then
            txtRemarks.SelText = " " & yClassForms.LastSearchName & " "
        Else
            txtRemarks.SelText = yClassForms.LastSearchName & " "
        End If
        txtRemarks.SelStart = Len(txtRemarks.Text)
    End If
End If
End Sub

Private Sub txtRemarks_LostFocus()
    txtRemarks.BackColor = &H80000005
End Sub

Private Sub txtRemarks2_GotFocus()
 txtRemarks2.BackColor = &HC0E0FF
End Sub



Private Sub txtRemarks2_LostFocus()
    txtRemarks2.BackColor = &H80000005
End Sub

Private Sub txtResultTag_Change()
If Len(Trim(lblProfileName.Caption)) < 1 Or Len(Trim(lblClassName.Caption)) < 1 Or getMandatoryFields = False Then
    cmdSavePreview.Enabled = False
    cmdSavePrint.Enabled = False
    cmdSave.Enabled = False
Else
    cmdSavePreview.Enabled = True
    cmdSavePrint.Enabled = True
    cmdSave.Enabled = True
End If
End Sub

Private Sub txtResultTag_DblClick()
    If GetSetting("HostForms_SA", "SelList", "cRoutineStat", "0") = 1 Then
        yClassForms.showSearchCode 11, GetSetting("HostForms_SA", "SelList", "RoutineStat", "Routine/Stat"), , lblResulttag.Caption, , True
        txtResultTag.Text = yClassForms.LastSearchName
    End If
End Sub

Private Sub txtResultTag_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtResultTag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        keybd_event VK_TAB, 0, 0, 0
    ElseIf KeyCode = vbKeyF2 Then
        If GetSetting("HostForms_SA", "SelList", "cRoutineStat", "0") = 1 Then
            yClassForms.showSearchCode 11, GetSetting("HostForms_SA", "SelList", "RoutineStat", "Routine/Stat"), , lblResulttag.Caption, , True
            txtResultTag.Text = yClassForms.LastSearchName
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If txtResultTag.Locked Then
            txtResultTag.Text = ""
        End If
    End If
End Sub

Private Sub txtResultTag_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtResultTag_LostFocus()
    txtResultTag.BackColor = &H80000005
End Sub

Private Sub txtRoom_Change()
If Len(Trim(lblProfileName.Caption)) < 1 Or Len(Trim(lblClassName.Caption)) < 1 Or getMandatoryFields = False Then
    cmdSavePreview.Enabled = False
    cmdSavePrint.Enabled = False
    cmdSave.Enabled = False
Else
    cmdSavePreview.Enabled = True
    cmdSavePrint.Enabled = True
    cmdSave.Enabled = True
End If
End Sub

Private Sub txtRoom_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtRoom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        keybd_event VK_TAB, 0, 0, 0
    ElseIf KeyCode = vbKeyF2 Then
        If GetSetting("HostForms_SA", "SelList", "cRoom", "0") = 1 Then
            yClassForms.showSearchCode 11, GetSetting("HostForms_SA", "SelList", "Room", "Room"), , lblRoom.Caption, , True
            txtRoom.Text = yClassForms.LastSearchName
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If txtRoom.Locked Then
            txtRoom.Text = ""
        End If
    End If
End Sub

Private Sub txtRoom_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtRoom_LostFocus()
    txtRoom.BackColor = &H80000005
End Sub

Private Sub txtSex_Change()
    If Val(GetSetting("HostForms_SA", "Settings", "DetectNV", "0")) = 1 Then
        If Len(Trim(txtSex.Text)) > 0 Then
            txtClassID.Text = getNVclass(txtSex.Text)
            If Len(Trim(txtClassID.Text)) > 0 Then
                yClassForms.showSearchCode 8, txtClassID, False
                txtClassID.Text = yClassForms.LastsearchID
                lblClassName.Caption = yClassForms.LastSearchName
                'If Len(Trim(lblClassName.Caption)) > 0 Then
            End If
        End If
    End If
    If Len(Trim(txtSex.Text)) > 0 Then
        If InStr(1, "MF", Left(txtSex.Text, 1)) > 0 Then
            txtSex_KeyPress Asc(Left(txtSex.Text, 1))
        End If
    End If
End Sub

Private Sub txtSex_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtSex_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtSex_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr(1, "MF", Chr(KeyAscii)) > 0 Then
        If Chr(KeyAscii) = "M" Then txtSex.Text = "MALE"
        If Chr(KeyAscii) = "F" Then txtSex.Text = "FEMALE"
    End If
    If KeyAscii = 32 Then
        If txtSex.Text = "MALE" Then
            txtSex.Text = "FEMALE"
        ElseIf txtSex.Text = "FEMALE" Then
            txtSex.Text = "MALE"
        Else
            txtSex.Text = "MALE"
        End If
    End If
    KeyAscii = 0
    txtSex_GotFocus
End Sub

Private Sub txtSex_LostFocus()
    txtSex.BackColor = &H80000005
End Sub

Private Sub txtSpecimen_Change()
If Len(Trim(lblProfileName.Caption)) < 1 Or Len(Trim(lblClassName.Caption)) < 1 Or getMandatoryFields = False Then
    cmdSavePreview.Enabled = False
    cmdSavePrint.Enabled = False
    cmdSave.Enabled = False
Else
    cmdSavePreview.Enabled = True
    cmdSavePrint.Enabled = True
    cmdSave.Enabled = True
End If
End Sub

Private Sub txtSpecimen_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtSpecimen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        keybd_event VK_TAB, 0, 0, 0
    ElseIf KeyCode = vbKeyF2 Then
        If GetSetting("HostForms_SA", "SelList", "cSpecimen", "0") = 1 Then
            yClassForms.showSearchCode 11, GetSetting("HostForms_SA", "SelList", "Specimen", "Specimen"), , lblSpecimen.Caption, , True
            txtSpecimen.Text = yClassForms.LastSearchName
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If txtSpecimen.Locked Then
            txtSpecimen.Text = ""
        End If
    End If
End Sub

Private Sub txtSpecimen_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtSpecimen_LostFocus()
    txtSpecimen.BackColor = &H80000005
End Sub

Private Sub txtSpecimenID_Change()
If Len(Trim(lblProfileName.Caption)) < 1 Or Len(Trim(lblClassName.Caption)) < 1 Or getMandatoryFields = False Then
    cmdSavePreview.Enabled = False
    cmdSavePrint.Enabled = False
    cmdSave.Enabled = False
Else
    cmdSavePreview.Enabled = True
    cmdSavePrint.Enabled = True
    cmdSave.Enabled = True
End If
End Sub

Private Sub txtSpecimenID_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtSpecimenID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then keybd_event VK_TAB, 0, 0, 0
End Sub

Private Sub txtSpecimenID_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtSpecimenID_LostFocus()
    txtSpecimenID.BackColor = &H80000005
End Sub

Private Sub txtWard_Change()
If Len(Trim(lblProfileName.Caption)) < 1 Or Len(Trim(lblClassName.Caption)) < 1 Or getMandatoryFields = False Then
    cmdSavePreview.Enabled = False
    cmdSavePrint.Enabled = False
    cmdSave.Enabled = False
Else
    cmdSavePreview.Enabled = True
    cmdSavePrint.Enabled = True
    cmdSave.Enabled = True
End If
End Sub

Private Sub txtWard_GotFocus()
    highLight
    highLight_bg
End Sub

Private Sub txtWard_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        keybd_event VK_TAB, 0, 0, 0
    ElseIf KeyCode = vbKeyF2 Then
        If GetSetting("HostForms_SA", "SelList", "cWard", "0") = 1 Then
            yClassForms.showSearchCode 11, GetSetting("HostForms_SA", "SelList", "Ward", "Ward"), , lblWard.Caption, , True
            txtWard.Text = yClassForms.LastSearchName
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If txtWard.Locked Then
            txtWard.Text = ""
        End If
    End If
End Sub

Private Sub txtWard_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub fillLists()
lstPhysician.Clear
lstMedtech.Clear
rstPhysician.Requery
If rstPhysician.State = 1 Then
    If rstPhysician.RecordCount > 0 Then rstPhysician.MoveFirst
    Do While Not rstPhysician.EOF
        lstPhysician.AddItem rstPhysician.Fields("drName").Value
        rstPhysician.MoveNext
    Loop
End If
rstMedtech.Requery
If rstMedtech.State = 1 Then
    If rstMedtech.RecordCount > 0 Then rstMedtech.MoveFirst
    Do While Not rstMedtech.EOF
        lstMedtech.AddItem rstMedtech.Fields("mtname").Value
        rstMedtech.MoveNext
    Loop
End If
End Sub

Public Property Let wMultiStart(ByVal vNewValue As Long)
    yMultiStart = vNewValue
End Property

Public Property Let wMultiTotal(ByVal vNewValue As Long)
    yMultiTotal = vNewValue
End Property

Public Property Let wMultiCurrent(ByVal vNewValue As Long)
    yMultiCurrent = vNewValue
End Property

Private Sub DisplayResultForm(Optional mShowResultEntry As Boolean = True)
Dim rstResultForm As ADODB.Recordset
Dim ctr As Long
msfgResult.Visible = False
If mShowResultEntry Then
    If txtProfileID.Text = "0" And Len(Trim(lblProfileName.Caption)) > 0 Then
        If NewResult And txtProfileID.Text = "0" Then
            txtProfileID.Text = ""
            If HasRefNum Then
                Set rstResultForm = myConnection.Execute("sp_DisplayExams '" & yInstrumentID & "','" & Replace(yPatientName, "'", "''") & "',#" & Format(yResultDate, "mm/dd/yyyy HH:mm:ss") & "#")
            Else
                Set rstResultForm = myConnection.Execute("sp_displayProfileExams '" & "DUMMY" & "','" & "DUMMY" & "'")
            End If
        Else
            Set rstResultForm = myConnection.Execute("sp_DisplayExams '" & yInstrumentID & "','" & Replace(yPatientName, "'", "''") & "',#" & Format(yResultDate, "mm/dd/yyyy HH:mm:ss") & "#")
        End If
    Else
        Set rstResultForm = myConnection.Execute("sp_displayProfileExams '" & txtProfileID.Text & "','" & InstrumentID & "'")
    End If
    If rstResultForm.RecordCount > 0 Then
        fraBody(2).ZOrder vbSendToBack
        Set msfgResult.DataSource = rstResultForm
        If Not NewResult Then setResults
        If NewResult And HasRefNum Then setResults
        If HasRefNum Then
            Set rstPatient2 = myConnection.Execute("Select * from tbPatientMaster where refnum = '" & txtRefNum & "'")
            txtClassID.Text = rstPatient2.Fields("ClassID").Value
            If Len(Trim(rstPatient2.Fields("ClassID").Value)) > 0 Then lblClassName.Caption = yClassForms.showSearchCode(2, txtClassID.Text)
        End If
        setSIvalues
        If Val(GetSetting("HostForms_SA", "Settings", "ReverseSI", "0")) = 1 Then
            msfgResult.TextMatrix(0, 3) = "Result (SI)"
            msfgResult.TextMatrix(0, 4) = "Normal Values (SI)"
            msfgResult.TextMatrix(0, 15) = "Result (C)"
            msfgResult.TextMatrix(0, 16) = "Normal Values (C)"
        End If
        lblClassName_Change
    Else
'        fraBody(1).Visible = False
        fraBody(1).ZOrder vbSendToBack
    End If
Else
'    fraBody(1).Visible = False
    fraBody(1).ZOrder vbSendToBack
End If
setExamRemarksCol
msfgResult.Visible = True
txtEntry.Visible = False
msfgResult.Row = 1
With msfgResult
    For ctr = 1 To .Rows - 1
        If .RowHeight(ctr) <> 0 Then .RowHeight(ctr) = 315
    Next ctr
End With
End Sub

Private Sub setResults()
Dim ctr As Integer
Dim tempStr As String
With msfgResult
.Redraw = False
'    If Not HasRefNum Then
        If DBServerType = "LOCAL DATABASE" Then
            Set rstClass = myConnection.Execute("Select * from tblabClassExams where Active = True and InstrumentID = '" & InstrumentID & "' and ClassID ='" & txtClassID.Text & "'")
        Else
            Set rstClass = myConnection.Execute("Select * from tblabClassExams where Active = 1 and InstrumentID = '" & InstrumentID & "' and ClassID ='" & txtClassID.Text & "'")
        End If
        For ctr = 1 To .Rows - 1
            rstPatient.MoveFirst
            rstPatient.Find "ExamName = '" & .TextMatrix(ctr, 6) & "'"
            If Not rstPatient.EOF Then
                .TextMatrix(ctr, 2) = rstPatient.Fields("Result").Value & ""
                If Not HasRefNum Then
                    tempStr = rstPatient.Fields("Result").Value & ""
                    If Len(Trim(tempStr)) > 0 And IsNumeric(tempStr) Then
                        tempStr = CDec(tempStr) * CDec(.TextMatrix(ctr, 7))
                        tempStr = Format(tempStr, IIf((Len(Trim(.TextMatrix(ctr, 8))) > 0), .TextMatrix(ctr, 8), ""))
                    Else
                        If Val(GetSetting("HostForms_SA", "Settings", "SaveStringResult", "0")) = 0 Then
                            tempStr = ""
                        End If
                    End If
                End If
                setNValuesItem ctr
                If .TextMatrix(ctr, 13) = "/" Then
                    .TextMatrix(ctr, 15) = Val(.TextMatrix(ctr, 3)) / Val(.TextMatrix(ctr, 17))
                    .TextMatrix(ctr, 15) = Format(.TextMatrix(ctr, 15), "0.00")
                Else
                    .TextMatrix(ctr, 15) = Val(.TextMatrix(ctr, 3)) * Val(.TextMatrix(ctr, 17))
                    .TextMatrix(ctr, 15) = Format(.TextMatrix(ctr, 15), "0.00")
                End If
                .TextMatrix(ctr, 3) = tempStr
            End If
        Next ctr
'    End If
.Redraw = True
End With
rstPatient.MoveFirst
setSavedValues
lblDiff.Caption = getTotalDiff
End Sub


Private Sub setNValues(Optional blankvalues As Boolean = False)
Dim ctr As Integer
If Not blankvalues Then
    If Len(Trim(lblProfileName.Caption)) > 0 Then
        With msfgResult
            For ctr = 1 To .Rows - 1
                setNValuesItem ctr
            Next ctr
        End With
    End If
Else
    With msfgResult
        For ctr = 1 To .Rows - 1
            .TextMatrix(ctr, 4) = ""
            .TextMatrix(ctr, 5) = ""
            .TextMatrix(ctr, 16) = ""
        Next ctr
    End With
End If
End Sub

Private Sub setNValuesItem(mRow As Integer)
Dim mHigh As String
Dim mLow As String
Dim mNormal As String
Dim mArray() As String
On Error GoTo errFound
mHigh = "High"
mLow = "Low"
mNormal = ""
With msfgResult
.Redraw = False
If Len(Trim(.TextMatrix(mRow, 20))) > 0 Then
    If InStr(1, .TextMatrix(mRow, 20), "|") Then
        mArray = Split(.TextMatrix(mRow, 20), "|")
        If UBound(mArray) >= 2 Then
            mHigh = mArray(2)
            mLow = mArray(0)
            mNormal = mArray(1)
        End If
    End If
End If
'If Not HasRefNum Then
    If rstClass.RecordCount > 0 Then
        rstClass.MoveFirst
        rstClass.Find "ExamID = '" & .TextMatrix(mRow, 9) & "'"
        If Not rstClass.EOF Then
            .TextMatrix(mRow, 4) = rstClass.Fields("Display").Value & ""
            .TextMatrix(mRow, 16) = rstClass.Fields("DisplaySI").Value & ""
            .TextMatrix(mRow, 12) = rstClass.Fields("RangeUnit").Value & ""
            .TextMatrix(mRow, 14) = rstClass.Fields("RangeUnitSI").Value & ""
            If Len(Trim(.TextMatrix(mRow, 3))) > 0 And IsNumeric(.TextMatrix(mRow, 3)) Then
                If CDbl(.TextMatrix(mRow, 3)) > Val(rstClass.Fields("RangeUp").Value) And CDbl(.TextMatrix(mRow, 3)) > Val(rstClass.Fields("RangeDown").Value) Then
                    .TextMatrix(mRow, 5) = mHigh
                ElseIf CDbl(.TextMatrix(mRow, 3)) < Val(rstClass.Fields("RangeDown").Value) And CDbl(.TextMatrix(mRow, 3)) < Val(rstClass.Fields("RangeUp").Value) Then
                    .TextMatrix(mRow, 5) = mLow
                Else
                    .TextMatrix(mRow, 5) = mNormal
                End If
            Else
                .TextMatrix(mRow, 5) = ""
            End If
            If Len(Trim(.TextMatrix(mRow, 3))) = 0 Then
                .TextMatrix(mRow, 15) = ""
            Else
                If Len(Trim(.TextMatrix(mRow, 3))) = 0 Then
                    .TextMatrix(mRow, 15) = ""
                Else
                    If .TextMatrix(mRow, 13) = "/" Then
                        .TextMatrix(mRow, 15) = Val(.TextMatrix(mRow, 3)) / Val(.TextMatrix(mRow, 17))
                        .TextMatrix(mRow, 15) = Format(.TextMatrix(mRow, 15), "0.00")
                    ElseIf .TextMatrix(mRow, 13) = "+" Then
                        .TextMatrix(mRow, 15) = Val(.TextMatrix(mRow, 3)) + Val(.TextMatrix(mRow, 17))
                        .TextMatrix(mRow, 15) = Format(.TextMatrix(mRow, 15), "0.00")
                    ElseIf .TextMatrix(mRow, 13) = "-" Then
                        .TextMatrix(mRow, 15) = Val(.TextMatrix(mRow, 3)) - Val(.TextMatrix(mRow, 17))
                        .TextMatrix(mRow, 15) = Format(.TextMatrix(mRow, 15), "0.00")
                    Else
                        .TextMatrix(mRow, 15) = Val(.TextMatrix(mRow, 3)) * Val(.TextMatrix(mRow, 17))
                        .TextMatrix(mRow, 15) = Format(.TextMatrix(mRow, 15), "0.00")
                    End If
                End If
            End If
        End If
    Else
        If Len(Trim(.TextMatrix(mRow, 3))) = 0 Then
            .TextMatrix(mRow, 15) = ""
        Else
            If .TextMatrix(mRow, 13) = "/" Then
                .TextMatrix(mRow, 15) = Val(.TextMatrix(mRow, 3)) / Val(.TextMatrix(mRow, 17))
                .TextMatrix(mRow, 15) = Format(.TextMatrix(mRow, 15), "0.00")
            ElseIf .TextMatrix(mRow, 13) = "+" Then
                .TextMatrix(mRow, 15) = Val(.TextMatrix(mRow, 3)) + Val(.TextMatrix(mRow, 17))
                .TextMatrix(mRow, 15) = Format(.TextMatrix(mRow, 15), "0.00")
            ElseIf .TextMatrix(mRow, 13) = "-" Then
                .TextMatrix(mRow, 15) = Val(.TextMatrix(mRow, 3)) - Val(.TextMatrix(mRow, 17))
                .TextMatrix(mRow, 15) = Format(.TextMatrix(mRow, 15), "0.00")
            Else
                .TextMatrix(mRow, 15) = Val(.TextMatrix(mRow, 3)) * Val(.TextMatrix(mRow, 17))
                .TextMatrix(mRow, 15) = Format(.TextMatrix(mRow, 15), "0.00")
            End If
        End If
    End If
'End If
.Redraw = True
End With
errFound:
    If Err.Number = 381 Then
    ElseIf Err.Number = 0 Then
    ElseIf Err.Number = 91 Then
        Resume Next
    Else
        MsgBox Err.Number & ": " & Err.Description
    End If
End Sub

Private Sub setSIvalues()
Dim ctr As Integer
On Error Resume Next
With msfgResult
For ctr = 1 To .Rows - 1
    If Len(Trim(.TextMatrix(ctr, 3))) = 0 Then
        .TextMatrix(ctr, 15) = ""
    Else
        If .TextMatrix(ctr, 13) = "/" Then
            .TextMatrix(ctr, 15) = Val(.TextMatrix(ctr, 3)) / Val(.TextMatrix(ctr, 17))
            .TextMatrix(ctr, 15) = Format(.TextMatrix(ctr, 15), "0.00")
        ElseIf .TextMatrix(ctr, 13) = "+" Then
            .TextMatrix(ctr, 15) = Val(.TextMatrix(ctr, 3)) + Val(.TextMatrix(ctr, 17))
            .TextMatrix(ctr, 15) = Format(.TextMatrix(ctr, 15), "0.00")
        ElseIf .TextMatrix(ctr, 13) = "-" Then
            .TextMatrix(ctr, 15) = Val(.TextMatrix(ctr, 3)) - Val(.TextMatrix(ctr, 17))
            .TextMatrix(ctr, 15) = Format(.TextMatrix(ctr, 15), "0.00")
        Else
            .TextMatrix(ctr, 15) = Val(.TextMatrix(ctr, 3)) * Val(.TextMatrix(ctr, 17))
            .TextMatrix(ctr, 15) = Format(.TextMatrix(ctr, 15), "0.00")
        End If
    End If
Next ctr
End With
End Sub

Private Sub SaveResult(Optional toPrint As Boolean = False, Optional SaveOnly As Boolean = False)
Dim rstTemp As ADODB.Recordset
Dim TimeStamp As String
Dim tempStr As String
Dim tempStatus As String
Dim ctr As Integer
Dim valDiff As Boolean
Dim strVerSaHdr As String
Dim strVerSaDtl As String
strVerSaHdr = ""
strVerSaDtl = ""
If Val(GetSetting("HostForms_SA", "Settings", "ByPassDiff", "0")) = 0 Then
    valDiff = validateDiff
Else
    If validateDiff = False Then
        MsgBox "Differential count is not equal to " & TotalDiff & ". Press Ok to continue.", vbInformation + vbOKOnly, "Diff Error"
    End If
    valDiff = True
End If
If valDiff Or SaveOnly Then
    fraBody(2).Enabled = False
    Me.MousePointer = 11
    TimeStamp = Now
    If HasRefNum = False Then
        HasRefNum = True
    End If
    'If Not NewResult Then myConnection.Execute "sp_UpdateInstrumentRefNum '" & yInstrumentID & "','" & Replace(yPatientName, "'", "''") & "',#" & Format(yResultDate, "mm/dd/yyyy HH:mm:ss") & "#,'" & txtRefNum.Text & "','" & Replace(txtSpecimenID.Text, "'", "''") & " ','" & Replace(txtLastName.Text, "'", "''") & ", " & Replace(txtFirstName.Text, "'", "''") & " " & Replace(txtMiddleName.Text, "'", "''") & "','1'"
    If yPatientName & yResultDate & ySequenceNumber <> "NEWNEWNEW" Then
        If DBServerType = "MSSQL SERVER" Then
            myConnection.Execute "sp_UpdateInstrumentRefNum  '" & yInstrumentID & "','" & Replace(yPatientName, "'", "''") & "','" & Format(yResultDate, "mm/dd/yyyy HH:mm:ss") & "','" & txtRefNum.Text & "','" & _
             Replace(txtSpecimenID.Text, "'", "''") & " ','" & Replace(txtLastName.Text, "'", "''") & ", " & Replace(txtFirstName.Text, "'", "''") & " " & Replace(txtMiddleName.Text, "'", "''") & "','1'"
        Else
            myConnection.Execute "sp_UpdateInstrumentRefNum  '" & yInstrumentID & "','" & Replace(yPatientName, "'", "''") & "',#" & Format(yResultDate, "mm/dd/yyyy HH:mm:ss") & "#,'" & txtRefNum.Text & "','" & Replace(txtSpecimenID.Text, "'", "''") & " ','" & Replace(txtLastName.Text, "'", "''") & ", " & Replace(txtFirstName.Text, "'", "''") & " " & Replace(txtMiddleName.Text, "'", "''") & "','1'"
        End If
    End If
    Set rstTemp = myConnection.Execute("sp_checkPatientData '" & txtRefNum.Text & "'")
    If rstTemp.RecordCount = 0 Then
        tempStr = "sp_InsertPatientData '" & txtRefNum.Text & "','" & Replace(txtLastName.Text, "'", "''") & "','" & Replace(txtFirstName.Text, "'", "''") & "','" & Replace(txtMiddleName.Text, "'", "''") & "','" & _
            Replace(txtAge.Text, "'", "''") & "','" & txtSex.Text & "','" & "1" & "','NV','','','" & Replace(txtBirthDate.Text, "'", "''") & "','" & Replace(txtWard.Text, "'", "''") & "','" & _
            Replace(txtRoom.Text, "'", "''") & "'," & returnDate(dtpResultDate.Value, dtpResultTime.Value) & "," & returnDate(dtpRequestDate.Value, dtpRequestTime.Value) & "," & _
            "NULL" & ",'" & Replace(txtPhysician.Text, "'", "''") & "','" & Replace(txtMedtech.Text, "'", "''") & "','" & txtProfileID.Text & "','" & txtClassID.Text & "'," & False & ",'" & InstrumentID & "','" & _
            Replace(yPatientName, "'", "''") & "'," & returnMemo(txtRemarks.Text) & ",'" & replaceStr(lblProfileName.Caption) & "','" & replaceStr(lblClassName.Caption) & "','" & replaceStr(txtPathologist.Text) & "','" & replaceStr(txtPatientID.Text) & "','" & replaceStr(txtSpecimenID.Text) & "','" & replaceStr(AdditionalProfiles) & "','" & _
            Replace(txtCivilStatus.Text, "'", "''") & "','" & Replace(txtSpecimen.Text, "'", "''") & "','" & Replace(txtResultTag.Text, "'", "''") & "','" & Replace(txtRemarks2.Text, "'", "''") & "'"
    Else
        tempStr = "sp_UpdatePatientData '" & txtRefNum.Text & "','" & replaceStr(txtLastName.Text) & "','" & replaceStr(txtFirstName.Text) & "','" & replaceStr(txtMiddleName.Text) & "','" & _
            replaceStr(txtAge.Text) & "','" & txtSex.Text & "','" & "1" & "','NV','','','" & Replace(txtBirthDate.Text, "'", "''") & "','" & replaceStr(txtWard.Text) & "','" & _
            replaceStr(txtRoom.Text) & "'," & returnDate(dtpResultDate.Value, dtpResultTime.Value) & "," & returnDate(dtpRequestDate.Value, dtpRequestTime.Value) & ",'" & _
            replaceStr(txtPhysician.Text) & "','" & replaceStr(txtMedtech.Text) & "','" & txtProfileID.Text & "','" & txtClassID.Text & "','" & InstrumentID & "','" & _
            replaceStr(yPatientName) & "'," & returnMemo(txtRemarks.Text) & ",'" & replaceStr(lblProfileName.Caption) & "','" & replaceStr(lblClassName.Caption) & "','" & replaceStr(txtPathologist.Text) & "','" & replaceStr(txtPatientID.Text) & "','" & replaceStr(txtSpecimenID.Text) & "','" & replaceStr(AdditionalProfiles) & "','" & _
            Replace(txtCivilStatus.Text, "'", "''") & "','" & Replace(txtSpecimen.Text, "'", "''") & "','" & Replace(txtResultTag.Text, "'", "''") & "','" & Replace(txtRemarks2.Text, "'", "''") & "'"
    End If
    strVerSaHdr = txtRefNum.Text & "|" & Replace(txtLastName.Text, "'", "''") & "|" & Replace(txtFirstName.Text, "'", "''") & "|" & Replace(txtMiddleName.Text, "'", "''") & "|" & _
            Replace(txtAge.Text, "'", "''") & "|" & txtSex.Text & "|" & Replace(txtBirthDate.Text, "'", "''") & "|" & Replace(txtWard.Text, "'", "''") & "|" & _
            Replace(txtRoom.Text, "'", "''") & "|" & Replace(returnDate(dtpResultDate.Value, dtpResultTime.Value), "#", "") & "|" & Replace(returnDate(dtpRequestDate.Value, dtpRequestTime.Value), "#", "") & "|" & _
            Replace(txtPhysician.Text, "'", "''") & "|" & Replace(txtMedtech.Text, "'", "''") & "|" & InstrumentID & "|" & _
            returnMemo(txtRemarks.Text) & "|" & replaceStr(lblProfileName.Caption) & "|" & replaceStr(lblClassName.Caption) & "|" & replaceStr(txtPathologist.Text) & "|" & replaceStr(txtPatientID.Text) & "|" & replaceStr(txtSpecimenID.Text) & "|" & replaceStr(AdditionalProfiles) & "|" & _
            Replace(txtCivilStatus.Text, "'", "''") & "|" & Replace(txtSpecimen.Text, "'", "''")
    myConnection.Execute tempStr
    saveMDMT
    With msfgResult
    myConnection.Execute "sp_ClearResultItems '" & txtRefNum.Text & "'"
        myConnection.BeginTrans
        For ctr = 1 To .Rows - 1
                If Len(Trim(.TextMatrix(ctr, 20))) > 0 Then
                    tempStatus = .TextMatrix(ctr, 5)
                Else
                    tempStatus = .TextMatrix(ctr, 5)
                End If
                tempStr = "sp_InsertResultItems '" & txtRefNum.Text & "','" & .TextMatrix(ctr, 9) & "','" & Replace(.TextMatrix(ctr, 2), "'", "''") & "','" & _
                    Replace(.TextMatrix(ctr, 3), "'", "''") & "','" & .TextMatrix(ctr, 4) & "','" & .TextMatrix(ctr, 12) & "','" & tempStatus & "','" & .TextMatrix(ctr, 0) & "','" & Replace(.TextMatrix(ctr, 1), "'", "''") & "'," & ctr & ",'" & .TextMatrix(ctr, 16) & "','" & .TextMatrix(ctr, 14) & "','" & .TextMatrix(ctr, 15) & "','" & .TextMatrix(ctr, 18) & "','" & .TextMatrix(ctr, 21) & "'"
                
                If Val(GetSetting("HostForms_SA", "Settings", "SaveBlankResult", "1")) = 1 Then
                    myConnection.Execute tempStr
                    strVerSaDtl = strVerSaDtl & txtRefNum.Text & "|" & .TextMatrix(ctr, 9) & "|" & .TextMatrix(ctr, 2) & "|" & _
                         Replace(.TextMatrix(ctr, 3), "'", "''") & "|" & .TextMatrix(ctr, 4) & "|" & .TextMatrix(ctr, 12) & "|" & tempStatus & "|" & .TextMatrix(ctr, 0) & "|" & .TextMatrix(ctr, 1) & "|" & ctr & "|" & .TextMatrix(ctr, 16) & "|" & .TextMatrix(ctr, 14) & "|" & .TextMatrix(ctr, 15) & "|" & .TextMatrix(ctr, 18) & "|" & .TextMatrix(ctr, 21) & "|" & .TextMatrix(ctr, 6) & "<SP>"
                Else
                    If Len(Trim(.TextMatrix(ctr, 3))) > 0 Then
                        myConnection.Execute tempStr
                        strVerSaDtl = strVerSaDtl & txtRefNum.Text & "|" & .TextMatrix(ctr, 9) & "|" & .TextMatrix(ctr, 2) & "|" & _
                             Replace(.TextMatrix(ctr, 3), "'", "''") & "|" & .TextMatrix(ctr, 4) & "|" & .TextMatrix(ctr, 12) & "|" & tempStatus & "|" & .TextMatrix(ctr, 0) & "|" & .TextMatrix(ctr, 1) & "|" & ctr & "|" & .TextMatrix(ctr, 16) & "|" & .TextMatrix(ctr, 14) & "|" & .TextMatrix(ctr, 15) & "|" & .TextMatrix(ctr, 18) & "|" & .TextMatrix(ctr, 21) & "|" & .TextMatrix(ctr, 6) & "<SP>"
                    End If
                End If
            Next ctr
            If EnableVSP And toPrint Then
                ObjVSP.ProcessVSP strVerSaHdr, strVerSaDtl
            End If
            strVerSaHdr = ""
            strVerSaDtl = ""
            myConnection.CommitTrans
    End With
    fraBody(2).Enabled = True
    Me.MousePointer = 0
Else
    MsgBox "Differential count is not equal to " & TotalDiff, vbInformation + vbOKOnly, "Diff Error"
End If
End Sub

Private Function returnDate(mDate As String, mTime As String) As String

If IsDate(mDate) And IsDate(mTime) Then
    If DBServerType = "LOCAL DATABASE" Then
        returnDate = "#" & Format(mDate, "mm/dd/yyyy ") & Format(mTime, "HH:mm:ss") & "#"
    Else
        returnDate = "'" & Format(mDate, "mm/dd/yyyy ") & Format(mTime, "HH:mm:ss") & "'"
    End If
Else
    returnDate = "NULL"
End If
End Function

Private Function returnMemo(tempStr As String) As String
If Len(Trim(tempStr)) > 0 Then
    returnMemo = "'" & Replace(tempStr, "'", "''") & "'"
Else
    returnMemo = "NULL"
End If
End Function

Private Sub setSavedValues()
Dim ctr As Integer
Set rstSavedResult = myConnection.Execute("select * from tbPatientValues where refnum = '" & txtRefNum.Text & "'")
If rstSavedResult.RecordCount > 0 Then
    With msfgResult
        For ctr = 1 To .Rows - 1
            rstSavedResult.MoveFirst
            rstSavedResult.Find "ExamID = '" & .TextMatrix(ctr, 9) & "'"
            If Not rstSavedResult.EOF Then
                If Len(Trim(rstSavedResult.Fields("Result").Value)) > 0 Then .TextMatrix(ctr, 3) = rstSavedResult.Fields("Result").Value
                If Len(Trim(rstSavedResult.Fields("NValuesDisplay").Value)) > 0 Then .TextMatrix(ctr, 4) = rstSavedResult.Fields("NValuesDisplay").Value
                If Len(Trim(rstSavedResult.Fields("Status").Value)) > 0 Then .TextMatrix(ctr, 5) = rstSavedResult.Fields("Status").Value
                .TextMatrix(ctr, 0) = rstSavedResult.Fields("Recheck").Value & ""
            Else
                .TextMatrix(ctr, 3) = ""
                .TextMatrix(ctr, 5) = ""
                .TextMatrix(ctr, 0) = ""
            End If
        Next ctr
    End With
End If
End Sub

Private Function getTotalDiff() As String
Dim ctr As Integer
Dim tempInt As Double
If hasDiff Then
    lblTextDiff.Visible = True
    lblDiff.Visible = True
    tempInt = 0
    With msfgResult
        For ctr = 1 To .Rows - 1
            If Format(.TextMatrix(ctr, 11), ">") = "TRUE" Then
                tempInt = tempInt + Val(.TextMatrix(ctr, 3))
            End If
        Next ctr
    End With
    getTotalDiff = Format(Trim(tempInt), GetSetting("LAB-i", "Settings", "DiffCountFormat", "0"))
Else
    lblTextDiff.Visible = False
    lblDiff.Visible = False
End If
End Function

Private Function hasDiff() As Boolean
Dim ctr As Integer
Dim tempBln As Boolean
tempBln = False
With msfgResult
If .Cols > 5 Then
    For ctr = 1 To .Rows - 1
        If Format(.TextMatrix(ctr, 11), ">") = "TRUE" Then
            tempBln = True
            Exit For
        End If
    Next ctr
End If
End With
hasDiff = tempBln
End Function

Private Function hasSIvalues() As Boolean
Dim ctr As Integer
Dim tempBln As Boolean
tempBln = False
With msfgResult
If .Cols > 5 Then
    For ctr = 1 To .Rows - 1
        If Len(Trim(.TextMatrix(ctr, 15))) > 0 Then
            tempBln = True
            Exit For
        End If
    Next ctr
End If
End With
hasSIvalues = tempBln
End Function
Private Function validateDiff() As Boolean
If hasDiff Then
        If Val(Format(TotalDiff, GetSetting("LAB-i", "Settings", "DiffCountFormat", "0"))) = Val(Format(getTotalDiff, GetSetting("LAB-i", "Settings", "DiffCountFormat", "0"))) Then
            validateDiff = True
        Else
            validateDiff = False
        End If
Else
    validateDiff = True
End If
End Function

Private Sub saveMDMT()
Dim rstTemp As ADODB.Recordset
Set rstTemp = myConnection.Execute("Select * from tbPhysician where drName = '" & Replace(txtPhysician.Text, "'", "''") & "'")
If rstTemp.RecordCount < 1 Then
    If Len(Trim(txtPhysician.Text)) > 0 Then myConnection.Execute "Insert into tbPhysician (drName) select '" & Replace(txtPhysician.Text, "'", "''") & "'"
End If
Set rstTemp = myConnection.Execute("Select * from tbmedtech where mtName = '" & Replace(txtMedtech.Text, "'", "''") & "'")
If rstTemp.RecordCount < 1 Then
    If Len(Trim(txtMedtech.Text)) > 0 Then myConnection.Execute "Insert into tbmedtech (mtName) select '" & Replace(txtMedtech.Text, "'", "''") & "'"
End If
If Val(GetSetting("HostForms_SA", "Settings", "SaveDefaultMT", "0")) = 1 Then
    SaveSetting "HostForms_SA", "Settings", "SaveDefaultMTName", txtMedtech.Text
End If
fillLists
Set rstTemp = Nothing
End Sub

Private Function replaceStr(tempStr As String) As String
    replaceStr = Replace(tempStr, "'", "''")
End Function

Private Sub setAdditionalProfiles(mValues As String)
Dim mArr As Variant
Dim ctr As Integer
Dim tempStr As String
Dim examfound As Boolean
examfound = False
If Len(mValues) > 0 Then
    mArr = Split(mValues, ",")
    tempStr = ""
    For ctr = 0 To UBound(mArr)
        tempStr = tempStr & "'" & mArr(ctr) & "',"
    Next ctr
    If Len(tempStr) > 0 Then
        tempStr = Mid(tempStr, 1, Len(tempStr) - 1)
        Set rstAdditionalProfiles = myConnection.Execute("select  b.profilename, a.* from tblabprofileexams a left outer join tblabprofiles b on a.profileid = b.profileid and a.instrumentid = b.instrumentid where a.profileid in (" & tempStr & ") and a.instrumentid ='" & InstrumentID & "'")
    End If
    With msfgResult
        For ctr = 1 To .Rows - 1
            rstAdditionalProfiles.MoveFirst
            Do While Not rstAdditionalProfiles.EOF
                If .TextMatrix(ctr, 9) = rstAdditionalProfiles.Fields("ExamID").Value Then
                    .TextMatrix(ctr, 18) = rstAdditionalProfiles.Fields("ProfileName").Value
                    examfound = True
                End If
                rstAdditionalProfiles.MoveNext
            Loop
            If Not examfound Then
                .Rows = .Rows + 1
            Else
                examfound = False
            End If
        Next ctr
    End With
End If
End Sub

Private Sub txtWard_LostFocus()
    txtWard.BackColor = &H80000005
End Sub

Private Function FormulaValue(mFormula As String) As String
Dim ctr As Integer
Dim tempFormula As String
Dim tempValue As String
Dim value1, value2 As Double
Dim getValue As Boolean
Dim valueCount As Integer
Dim tempFunction As String
Dim myValue As Double
Dim isConstant As Boolean
Dim hasBlank As Boolean
Dim tempResultValue As String
hasBlank = False
value1 = 0
value2 = 0
tempValue = ""
valueCount = 0
getValue = False
isConstant = False
myValue = 0
If InStr(1, mFormula, "|") Then

Else
    For ctr = 1 To Len(Trim(mFormula))
        Select Case Mid(mFormula, ctr, 1)
            Case "<"
                getValue = True
                valueCount = valueCount + 1
            Case ">"
                If getValue Then
                    getValue = False
                    If valueCount = 2 Then
                        tempResultValue = getResultValue(tempValue)
                        If Len(Trim(tempResultValue)) = 0 Then hasBlank = True
                        value2 = Val(tempResultValue)
                        Select Case tempFunction
                            Case "+"
                                myValue = value1 + value2
                            Case "-"
                                myValue = value1 - value2
                            Case "*"
                                myValue = value1 * value2
                            Case "/"
                                If value1 = 0 Or value2 = 0 Then
                                    myValue = 0
                                Else
                                    myValue = value1 / value2
                                End If
                        End Select
                        valueCount = 1
                        value1 = myValue
                        tempFunction = ""
                        tempValue = ""
                    Else
                        tempResultValue = getResultValue(tempValue)
                        If Len(Trim(tempResultValue)) = 0 Then hasBlank = True
                        value1 = Val(tempResultValue)
                        tempValue = ""
                    End If
                End If
            Case "("
            
            Case ")"
            Case "-", "+", "/", "*", "^"
                tempFunction = Mid(mFormula, ctr, 1)
            Case " "
                If IsNumeric(tempValue) And Not getValue Then
                    If valueCount = 1 Then
                        value2 = Val(tempValue)
                        Select Case tempFunction
                            Case "+"
                                myValue = value1 + value2
                            Case "-"
                                myValue = value1 - value2
                            Case "*"
                                myValue = value1 * value2
                            Case "/"
                                If value1 = 0 Or value2 = 0 Then
                                    myValue = 0
                                Else
                                    myValue = value1 / value2
                                End If
                        End Select
                        valueCount = 1
                        value1 = myValue
                        tempFunction = ""
                    Else
                        tempResultValue = getResultValue(tempValue)
                        If Len(Trim(tempResultValue)) = 0 Then hasBlank = True
                        value1 = Val(tempResultValue)
                        tempValue = ""
                    End If
                End If
            Case Else
                tempValue = tempValue & Mid(mFormula, ctr, 1)
                If IsNumeric(tempValue) And Not getValue And ctr = Len(mFormula) Then
                    If valueCount = 1 Then
                        value2 = Val(tempValue)
                        Select Case tempFunction
                            Case "+"
                                myValue = value1 + value2
                            Case "-"
                                myValue = value1 - value2
                            Case "*"
                                myValue = value1 * value2
                            Case "/"
                                If value1 = 0 Or value2 = 0 Then
                                    myValue = 0
                                Else
                                    myValue = value1 / value2
                                End If
                        End Select
                        valueCount = 0
                        value1 = myValue
                        tempFunction = ""
                    Else
                        value1 = Val(tempValue)
                        tempValue = ""
                    End If
                End If
        End Select
'        tempFormula = Mid(mFormula, ctr, 1)
    Next ctr
End If
If hasBlank Then
    FormulaValue = ""
Else
    If InStr(1, mFormula, "+") > 0 Or InStr(1, mFormula, "-") > 0 Or InStr(1, mFormula, "*") > 0 Or InStr(1, mFormula, "/") > 0 Or InStr(1, mFormula, "^") > 0 Then
        FormulaValue = myValue
    Else
        FormulaValue = value1 + value2
    End If
End If
End Function

Private Function getResultValue(mCode As String) As String
Dim ctr As Integer
getResultValue = "0"
With msfgResult
    For ctr = 1 To .Rows - 1
        If Trim(.TextMatrix(ctr, 9)) = Trim(mCode) Then
            getResultValue = .TextMatrix(ctr, 3)
            Exit For
        Else
            getResultValue = ""
        End If
    Next ctr
End With
End Function

Private Sub setComputedValue()
Dim tempStr As String
Dim ctr As Integer
On Error Resume Next
With msfgResult
    For ctr = 1 To .Rows - 1
        If Len(Trim(.TextMatrix(ctr, 19))) > 0 Then
            tempStr = FormulaValue(.TextMatrix(ctr, 19))
            .TextMatrix(ctr, 3) = Format(tempStr, .TextMatrix(ctr, 8))
        End If
    Next ctr
End With
End Sub

Public Property Let wProfileID(ByVal vNewValue As String)
    yProfileID = vNewValue
End Property

Private Function GetIdFromName(PName As String) As String
Dim ctr As Integer
Dim tempStr As String
Dim tempStr2 As String
Dim hasNumeric As Boolean
Dim hasAlpha As Boolean
hasNumeric = False
hasAlpha = False
tempStr = ""
tempStr2 = ""
For ctr = 1 To Len(PName)
    If IsNumeric(Mid(PName, ctr, 1)) Then
        If Not hasAlpha Then
            tempStr = tempStr & Mid(PName, ctr, 1)
            hasNumeric = True
        End If
    Else
        If hasNumeric Then
            tempStr2 = tempStr2 & Mid(PName, ctr, 1)
            hasAlpha = True
        Else
            tempStr2 = tempStr2 & Mid(PName, ctr, 1)
        End If
    End If
Next ctr
GetIdFromName = tempStr & "|" & tempStr2
End Function

Private Function getUpdatedValues(mRefNum As String) As String
Dim rstTemp As ADODB.Recordset
Set rstTemp = myConnection.Execute("Select * from tbLabInstrumentResults where refnum = '" & mRefNum & "'")
If rstTemp.RecordCount > 0 Then
    getUpdatedValues = rstTemp.Fields("resultdate").Value & "|" & rstTemp.Fields("NewPatientName").Value & "|" & rstTemp.Fields("RefNum").Value & "|" & rstTemp.Fields("NewPatientName").Value
Else
    getUpdatedValues = ""
End If
End Function

Private Sub setPosDemo()
DoEvents
fraBody(4).Width = fraBody(0).Width / 3
fraBody(5).Width = fraBody(0).Width / 3
fraBody(6).Width = fraBody(0).Width / 3
fraBody(4).Left = (fraBody(0).Width / 2) - (fraBody(4).Width / 2)
fraBody(5).Left = fraBody(0).Width - fraBody(5).Width
'~~~~~~~~~~~ 1st Column ~~~~~~~~~~~
txtPatientID.Width = fraBody(6).Width - txtPatientID.Left
txtLastName.Width = txtPatientID.Width
txtBirthDate.Width = txtPatientID.Width
Label3.Left = fraBody(6).Width - Label3.Width
txtPhysician.Width = txtPatientID.Width - Label3.Width
lblProfileidB.Left = Label3.Left
lblProfileName.Width = txtPatientID.Width - txtProfileID.Width - Label3.Width
txtAge.Width = ((txtPatientID.Width - lblSex.Width) / 2) - 100
txtSex.Width = txtAge.Width
txtSex.Left = fraBody(6).Width - txtSex.Width
lblSex.Left = txtAge.Left + txtAge.Width + 100
'~~~~~~~~~~~ 2nd Column ~~~~~~~~~~~
txtSpecimenID.Width = fraBody(4).Width - txtSpecimenID.Left
txtFirstName.Width = txtSpecimenID.Width
txtCivilStatus.Width = txtSpecimenID.Width
Label2.Left = fraBody(4).Width - Label2.Width
txtMedtech.Width = txtSpecimenID.Width - Label2.Width
lblClassB.Left = Label2.Left
lblClassName.Width = txtSpecimenID.Width - Label2.Width - txtClassID.Width
txtWard.Width = ((txtSpecimenID.Width - lblRoom.Width) / 2) - 100
txtRoom.Width = txtWard.Width
txtRoom.Left = fraBody(6).Width - txtRoom.Width
lblRoom.Left = txtWard.Left + txtWard.Width + 100
'~~~~~~~~~~~ 3rd Column ~~~~~~~~~~~
lblText(0).Left = fraBody(5).Left
txtRefNum.Left = fraBody(5).Left + lblText(0).Width + 100
txtRefNum.Width = txtPatientID.Width - 300
txtMiddleName.Width = txtPatientID.Width - 300
txtSpecimen.Width = txtPatientID.Width - 300
txtResultTag.Width = txtPatientID.Width - 300
Label1.Left = (fraBody(5).Width - Label1.Width) - 200
txtPathologist.Width = txtPatientID.Width - Label1.Width - Label1.Width - 100
dtpResultDate.Width = (txtPatientID.Width / 2) - 300
dtpRequestDate.Width = dtpResultDate.Width
dtpResultTime.Width = dtpResultDate.Width
dtpRequestTime.Width = dtpResultDate.Width
dtpResultTime.Left = dtpResultDate.Left + dtpResultDate.Width + 200
dtpRequestTime.Left = dtpResultTime.Left
End Sub

Private Sub setExamRemarksCol()
Dim ctr As Integer
On Error Resume Next
With msfgResult
    For ctr = 1 To .Rows - 1
        If Len(Trim(.TextMatrix(ctr, 0))) > 0 Then
            .Col = 0
            .Row = ctr
            .CellBackColor = .CellForeColor
            If .TextMatrix(ctr, 22) = "SV" Then
                .Col = 3
'                .CellBackColor = &H8000000F
'                .CellForeColor = &H8000000F
                .CellBackColor = &H8000000F
                .CellForeColor = .ForeColor
            End If
        Else
            If .Cols >= 23 Then
                If .TextMatrix(ctr, 22) = "SV" Then
                    .Row = ctr
                    .Col = 3
                    .CellBackColor = &H8000000F
                    .CellForeColor = &H8000000F
                Else
                    .Col = 0
                    .Row = ctr
                    .CellBackColor = .BackColorFixed
                End If
            Else
                .Col = 0
                .Row = ctr
                .CellBackColor = .BackColorFixed
            End If
        End If
    Next ctr
End With
End Sub

Private Function getMandatoryFields() As Boolean
Dim enableSave As Boolean
enableSave = True
    If Val(GetSetting("HostForms_SA", "Settings", "ManFld01", "0")) = 1 And Len(Trim(txtPatientID.Text)) = 0 Then
        enableSave = False
        GoTo finalStat
    End If
    
    If Val(GetSetting("HostForms_SA", "Settings", "ManFld02", "0")) = 1 And Len(Trim(txtSpecimenID.Text)) = 0 Then
        enableSave = False
        GoTo finalStat
    End If
    
    If Val(GetSetting("HostForms_SA", "Settings", "ManFld03", "0")) = 1 And Len(Trim(txtWard.Text)) = 0 Then
        enableSave = False
        GoTo finalStat
    End If
    
    If Val(GetSetting("HostForms_SA", "Settings", "ManFld04", "0")) = 1 And Len(Trim(txtRoom.Text)) = 0 Then
        enableSave = False
        GoTo finalStat
    End If
    
    If Val(GetSetting("HostForms_SA", "Settings", "ManFld05", "0")) = 1 And Len(Trim(txtCivilStatus.Text)) = 0 Then
        enableSave = False
        GoTo finalStat
    End If
    
    If Val(GetSetting("HostForms_SA", "Settings", "ManFld06", "0")) = 1 And Len(Trim(txtSpecimen.Text)) = 0 Then
        enableSave = False
        GoTo finalStat
    End If
    
    If Val(GetSetting("HostForms_SA", "Settings", "ManFld07", "0")) = 1 And Len(Trim(txtResultTag.Text)) = 0 Then
        enableSave = False
        GoTo finalStat
    End If
finalStat:
    getMandatoryFields = enableSave
End Function

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

Private Function countVisibleGrid() As Integer
Dim ctr As Integer
Dim mycount As Integer
With msfgResult
    mycount = 0
    For ctr = 1 To .Rows - 1
        If .RowHeight(ctr) > 0 Then mycount = mycount + 1
    Next ctr
    countVisibleGrid = mycount
End With
End Function

'Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
'  Dim ctl As Control
'
'  For Each ctl In Me.Controls
'    If TypeOf ctl Is MSHFlexGrid Then
'      If IsOver(ctl.hWnd, Xpos, Ypos) Then FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
'    End If
'  Next ctl
'End Sub
