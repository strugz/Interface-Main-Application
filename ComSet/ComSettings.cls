VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function showSettings()
With frmSettings2
    .Show 1
End With
End Function

Public Property Get Settings(Optional ctrlnum As String = "PortSettings") As String
Settings = GetSetting("ComSet", ctrlnum, "BaudRate", "9600") & "," & _
    Mid(GetSetting("ComSet", ctrlnum, "Parity", "NONE"), 1, 1) & "," & _
    GetSetting("ComSet", ctrlnum, "DataBits", "8") & "," & _
    GetSetting("ComSet", ctrlnum, "StopBits", "1")
End Property

Public Property Get ComPort(Optional ctrlnum As String = "PortSettings") As String
ComPort = Mid(GetSetting("ComSet", ctrlnum, "ComPort", "COM1"), 4, 1)
End Property

Public Property Get HandShaking(Optional ctrlnum As String = "PortSettings") As String
HandShaking = Mid(GetSetting("ComSet", ctrlnum, "HandShake", "0-ComNone"), 1, 1)
End Property

Public Property Get InBufferSize(Optional ctrlnum As String = "PortSettings") As String
InBufferSize = GetSetting("ComSet", ctrlnum, "InbufferSize", "1024")
End Property

Public Property Get OutBufferSize(Optional ctrlnum As String = "PortSettings") As String
OutBufferSize = GetSetting("ComSet", ctrlnum, "OutbufferSize", "512")
End Property

Public Property Get RTresHold(Optional ctrlnum As String = "PortSettings") As String
RTresHold = GetSetting("ComSet", ctrlnum, "RTresHold", "1")
End Property

Public Property Get STresHold(Optional ctrlnum As String = "PortSettings") As String
STresHold = GetSetting("ComSet", ctrlnum, "STresHold", "1")
End Property

Public Property Get InputLen(Optional ctrlnum As String = "PortSettings") As String
InputLen = GetSetting("ComSet", ctrlnum, "InputLen", "0")
End Property

Public Property Get ParityReplace(Optional ctrlnum As String = "PortSettings") As String
ParityReplace = GetSetting("ComSet", ctrlnum, "ParityReplace", "?")
End Property

Public Property Get EOFEnable(Optional ctrlnum As String = "PortSettings") As Boolean
EOFEnable = IIf(Val(GetSetting("ComSet", ctrlnum, "EOFEnable", "0")) = "1", True, False)
End Property

Public Property Get NullDiscard(Optional ctrlnum As String = "PortSettings") As Boolean
NullDiscard = IIf(Val(GetSetting("ComSet", ctrlnum, "NullDiscard", "1")) = "1", True, False)
End Property

Public Property Get RTSEnable(Optional ctrlnum As String = "PortSettings") As Boolean
RTSEnable = IIf(Val(GetSetting("ComSet", ctrlnum, "RTSEnable", "1")) = "1", True, False)
End Property

Public Property Get DTREnable(Optional ctrlnum As String = "PortSettings") As Boolean
DTREnable = IIf(Val(GetSetting("ComSet", ctrlnum, "DTREnable", "1")) = "1", True, False)
End Property

Public Property Get ControlEnabled(Optional ctrlnum As String = "PortSettings") As String
ControlEnabled = Val(GetSetting("ComSet", ctrlnum, "Enabled", "0"))
End Property
