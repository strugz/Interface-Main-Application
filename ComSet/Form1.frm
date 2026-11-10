VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   150
      TabIndex        =   2
      Top             =   4350
      Width           =   1245
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4125
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   7276
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   585
      Left            =   4200
      TabIndex        =   0
      Top             =   4350
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private settings As New ComSettings
Private cforms As New ClassForms
Private varTemp As InstrumentValues

Private Sub Command1_Click()
'    MsgBox getSTKsData(RichTextBox1.Text, "NAME", "NO NAME", Val(Text1.Text))
'    MsgBox settings.ComPort
'    MsgBox settings.settings
'    MsgBox settings.HandShaking
'    MsgBox settings.inbuffersize
'    MsgBox settings.outbuffersize
'    MsgBox settings.rtreshold
'    MsgBox settings.streshold
'    MsgBox settings.inputlen
'    MsgBox settings.eofenable
'    MsgBox settings.parityreplace
'    MsgBox settings.nulldiscard
'    MsgBox settings.rtsenable
'    MsgBox settings.dtrenable
'    varTemp = cforms.ShowSelectInstruments("1")
'    MsgBox varTemp.InstrumentID
'    MsgBox varTemp.InstrumentName
'    MsgBox varTemp.LabSectionCode
'    MsgBox varTemp.Plugin
'    MsgBox varTemp.settings
    settings.showsettings
    
End Sub

Private Sub Form_Load()
    'Set settings = CreateObject("ComSet.ComSettings")
End Sub
