VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowResultViewer(mInstrumentID As String, mPatientName As String, mDateTime As String)
With frmResultViewer
    .wInstrumentID = mInstrumentID
    .wPatientName = mPatientName
    .wDateTime = mDateTime
    .Show 1
End With
End Sub

Public Function showInputBox(mTitle As String, Optional mTop As Integer, Optional mLeft As Integer, Optional mValue As String = "", Optional mMaxLength As Integer = 0, Optional mUpCase As Boolean = False, Optional Mask As Boolean = False) As String
With frmInputBox
    If mTop <> 0 Then .wTop = mTop
    If mLeft <> 0 Then .wLeft = mLeft
    .wTitle = mTitle
    .wValue = mValue
    .wMaxLength = mMaxLength
    .wUpcase = mUpCase
    .wMask = Mask
    .Show 1
    showInputBox = frmSearchVal
End With
End Function

Public Function showArchive() As String
ReturnValue = ""
frmArchive.Show 1
If ReturnValue = "" Then
    showArchive = ""
Else
    showArchive = ReturnValue
End If
End Function

