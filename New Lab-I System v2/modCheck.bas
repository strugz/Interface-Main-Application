Attribute VB_Name = "modCheck"
Option Explicit
Private OnBits(0 To 31) As Long

Private Const Table_Lookup = "&H0000,&H1021,&H2042,&H3063,&H4084,&H50A5,&H60C6,&H70E7,&H8108,&H9129,&HA14A,&HB16B,&HC18C,&HD1AD,&HE1CE,&HF1EF,&H1231,&H0210,&H3273,&H2252,&H52B5,&H4294,&H72F7,&H62D6," & _
"&H9339,&H8318,&HB37B,&HA35A,&HD3BD,&HC39C,&HF3FF,&HE3DE,&H2462,&H3443,&H0420,&H1401,&H64E6,&H74C7,&H44A4,&H5485,&HA56A,&HB54B,&H8528,&H9509,&HE5EE,&HF5CF,&HC5AC,&HD58D," & _
"&H3653,&H2672,&H1611,&H0630,&H76D7,&H66F6,&H5695,&H46B4,&HB75B,&HA77A,&H9719,&H8738,&HF7DF,&HE7FE,&HD79D,&HC7BC,&H48C4,&H58E5,&H6886,&H78A7,&H0840,&H1861,&H2802,&H3823," & _
"&HC9CC,&HD9ED,&HE98E,&HF9AF,&H8948,&H9969,&HA90A,&HB92B,&H5AF5,&H4AD4,&H7AB7,&H6A96,&H1A71,&H0A50,&H3A33,&H2A12,&HDBFD,&HCBDC,&HFBBF,&HEB9E,&H9B79,&H8B58,&HBB3B,&HAB1A," & _
"&H6CA6,&H7C87,&H4CE4,&H5CC5,&H2C22,&H3C03,&H0C60,&H1C41,&HEDAE,&HFD8F,&HCDEC,&HDDCD,&HAD2A,&HBD0B,&H8D68,&H9D49,&H7E97,&H6EB6,&H5ED5,&H4EF4,&H3E13,&H2E32,&H1E51,&H0E70," & _
"&HFF9F,&HEFBE,&HDFDD,&HCFFC,&HBF1B,&HAF3A,&H9F59,&H8F78,&H9188,&H81A9,&HB1CA,&HA1EB,&HD10C,&HC12D,&HF14E,&HE16F,&H1080,&H00A1,&H30C2,&H20E3,&H5004,&H4025,&H7046,&H6067," & _
"&H83B9,&H9398,&HA3FB,&HB3DA,&HC33D,&HD31C,&HE37F,&HF35E,&H02B1,&H1290,&H22F3,&H32D2,&H4235,&H5214,&H6277,&H7256,&HB5EA,&HA5CB,&H95A8,&H8589,&HF56E,&HE54F,&HD52C,&HC50D," & _
"&H34E2,&H24C3,&H14A0,&H0481,&H7466,&H6447,&H5424,&H4405,&HA7DB,&HB7FA,&H8799,&H97B8,&HE75F,&HF77E,&HC71D,&HD73C,&H26D3,&H36F2,&H0691,&H16B0,&H6657,&H7676,&H4615,&H5634," & _
"&HD94C,&HC96D,&HF90E,&HE92F,&H99C8,&H89E9,&HB98A,&HA9AB,&H5844,&H4865,&H7806,&H6827,&H18C0,&H08E1,&H3882,&H28A3,&HCB7D,&HDB5C,&HEB3F,&HFB1E,&H8BF9,&H9BD8,&HABBB,&HBB9A," & _
"&H4A75,&H5A54,&H6A37,&H7A16,&H0AF1,&H1AD0,&H2AB3,&H3A92,&HFD2E,&HED0F,&HDD6C,&HCD4D,&HBDAA,&HAD8B,&H9DE8,&H8DC9,&H7C26,&H6C07,&H5C64,&H4C45,&H3CA2,&H2C83,&H1CE0,&H0CC1," & _
"&HEF1F,&HFF3E,&HCF5D,&HDF7C,&HAF9B,&HBFBA,&H8FD9,&H9FF8,&H6E17,&H7E36,&H4E55,&H5E74,&H2E93,&H3EB2,&H0ED1,&H1EF0"


Public Function sumChr(myMess As String) As Integer
Dim ctr As Integer
sumChr = 0
For ctr = 1 To Len(myMess)
    sumChr = sumChr + Asc(Mid(myMess, ctr, 1))
Next ctr
End Function

Public Function dec2bin(mynum As Long) As String
Dim myDump As Long
Dim myCtr1 As Long
myCtr1 = mynum
dec2bin = ""
Do While myCtr1 > 0
     myDump = myCtr1 Mod 2
      If myDump = 1 Then
        myCtr1 = (myCtr1 - 1) / 2
    Else
        myCtr1 = myCtr1 / 2
     End If
    dec2bin = Trim(Str(myDump)) & Trim(dec2bin)
Loop
If Len(dec2bin) Mod 4 <> 0 Then
    dec2bin = String(4 - Len(dec2bin) Mod 4, "0") & dec2bin
End If
End Function

Public Function bin2dec(myBinString As String) As Long
Dim ctr As Long
Dim myArray As Long
Dim myDump As Long
bin2dec = 0
For ctr = Len(myBinString) To 1 Step -1
    myDump = Mid(myBinString, ctr, 1)
    If myDump = "1" Then myArray = 2 ^ (Len(myBinString) - ctr) Else myArray = 0
    bin2dec = bin2dec + myArray
Next ctr
End Function

Public Function CheckSum8(myMessage As String) As String
    CheckSum8 = Hex$(bin2dec(dec2bin(sumChr(myMessage))))
    If Len(CheckSum8) = 1 Then CheckSum8 = "0" & CheckSum8
End Function

Public Function isValidCS(myMessage As String) As Boolean
Dim strStart As Integer
Dim strEnd As Integer
Dim tempStr As String
Dim tempChecksum As String
On Error GoTo errFound
    strStart = InStr(1, myMessage, Chr(2)) + 1
    strEnd = InStr(1, myMessage, Chr(3))
    tempStr = Mid(myMessage, strStart, (strEnd - strStart) - 2)
    tempChecksum = Mid(myMessage, strEnd - strStart, 2)
    If CheckSum8(tempStr) = tempChecksum Then isValidCS = True Else isValidCS = False
errFound:
    If Err = 5 Then isValidCS = False
End Function


Public Function MsgCRC(Message As String) As String
'Calculate CRC checksum for message then add
'to end of message. Check for 10h byte 'DLE'
'within message and change to 'DLE DLE'.


'Function to calculate message CRC

Dim xByte As Integer
Dim xResult As Long

xByte = 3

Do

xResult = xResult Xor Asc(Mid(Message, xByte, 1))

RotateResult xResult
'test for 'DLE' and double up to DLE DLE
If Asc(Mid(Message, xByte, 1)) = 16 Then
'add extra DLE
Message = Left$(Message, xByte) + Chr(16) + Right$(Message, Len(Message) - xByte)

xByte = xByte + 1 'don't include in CRC

End If
xByte = xByte + 1

Loop While (xByte <= Len(Message) - 2)

xResult = xResult Xor 3 'ETX byte
RotateResult xResult

'add CRC checksum to message
MsgCRC = Chr(xResult Mod 256) + Chr(Int(xResult / 256))

End Function

Private Sub RotateResult(res&)
Dim bitout%, Shift%

'Rotate result 8 times and combine with const.
For Shift% = 1 To 8
bitout% = res& Mod 2 'test if bit will be shifted out
res& = Int(res& / 2) 'shift right
If bitout% Then
res& = res& Xor &H1000A001 'xor with constant
res& = res& - &H10000000 'clear top word
End If
Next Shift%

End Sub


Public Function CRC16_CCITT_M(myStr As String) As String
Dim mValue As Long
Dim ctr As Integer
Dim X As Integer
Dim LTable() As String
LTable = Split(Table_Lookup, ",")
mValue = 65535
For ctr = 1 To Len(myStr)
    X = RShiftLong(CStr(mValue), 8) Xor Asc(Mid(myStr, ctr, 1))
    mValue = (LShiftLong(CStr(mValue), 8) And 65535) Xor CDec(LTable(X))
Next ctr
CRC16_CCITT_M = Right(Hex$(mValue Xor &HFFFF), 4)
End Function

Public Function CRC16_CCITT(myStr As String) As String
Dim mValue As Long
Dim ctr As Integer
Dim X As Integer
Dim LTable() As String
LTable = Split(Table_Lookup, ",")
mValue = 65535
For ctr = 1 To Len(myStr)
    X = RShiftLong(CStr(mValue), 8) Xor Asc(Mid(myStr, ctr, 1))
    mValue = (LShiftLong(CStr(mValue), 8) And 65535) Xor CDec(LTable(X))
Next ctr
CRC16_CCITT = Hex$(mValue)
End Function


Private Function LShiftLong(ByVal Value As Long, ByVal Shift As Integer) As Long
MakeOnBits
If (Value And (2 ^ (31 - Shift))) Then GoTo OverFlow
LShiftLong = ((Value) * (2 ^ Shift))
Exit Function
OverFlow:
    LShiftLong = ((Value And OnBits(31 - (Shift + 1))) * _
       (2 ^ (Shift))) Or &H80000000
End Function

Private Function RShiftLong(ByVal Value As Long, _
   ByVal Shift As Integer) As Long
    Dim hi As Long
    MakeOnBits
    If (Value And &H80000000) Then hi = &H40000000
    RShiftLong = (Value And &H7FFFFFFE) \ (2 ^ Shift)
    RShiftLong = (RShiftLong Or (hi \ (2 ^ (Shift - 1))))
End Function
 
Private Sub MakeOnBits()
Dim j As Integer, _
    v As Long
    For j = 0 To 30
        v = v + (2 ^ j)
        OnBits(j) = v
    Next j
    OnBits(j) = v + &H80000000
End Sub

Public Function ASTMChecksum(myStr As String) As String
Dim checksum As Long
Dim lenCtr As Integer
checksum = 0
lenCtr = 1
Do While lenCtr <= Len(myStr)
    checksum = checksum + Asc(Mid(myStr, lenCtr, 1)) And 255
    lenCtr = lenCtr + 1
Loop
ASTMChecksum = Right(Hex$(Val(checksum) + 16), 2)
End Function

