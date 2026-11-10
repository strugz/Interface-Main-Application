Attribute VB_Name = "modCRC"
Option Explicit

Public Function CRC16_CCITT2(mystring As String) As String
Dim lenStr As String
Dim initValue As Long
Dim ctr As Integer
Dim mCH As Long
Dim Poly As Integer
Dim XorFlag As Boolean
lenStr = 1
initValue = 65535
XorFlag = False
Poly = 4129
Do While lenStr <= Len(mystring)
    mCH = Asc(Mid(mystring, lenStr, 1))
    mCH = LShift2(mCH, 8)
    For ctr = 1 To 8
        If ((initValue Xor mCH) And 32768) Then
            XorFlag = True
         Else
            XorFlag = False
        End If
        initValue = LShift(initValue, 1)
        If XorFlag Then
            initValue = initValue Xor Poly
        End If
        mCH = LShift(mCH, 1)
    Next ctr
    lenStr = lenStr + 1
Loop
CRC16_CCITT2 = Hex$(initValue)
End Function

Public Function CRC16_CCITT(mystring As String) As String
Dim initValue As Long
Dim v As Long
Dim lenStr As Long
Dim ctr As Integer
Dim XorFlag As Boolean
Dim mCH As Integer
Dim Poly As Integer
XorFlag = False
initValue = 65535
lenStr = 1
Poly = 4129
Do While lenStr <= Len(mystring)
    v = 128
    mCH = Asc(Mid(mystring, lenStr, 1))
    For ctr = 1 To 8
        If initValue And 32768 Then
            XorFlag = True
        Else
            XorFlag = False
        End If
        initValue = LShift(initValue, 1)
        If mCH And v Then
            initValue = initValue + 1
        End If
        If XorFlag Then
            initValue = initValue Xor Poly
        End If
        v = RShift(v, 1)
    Next ctr
    lenStr = lenStr + 1
Loop
CRC16_CCITT = Hex$(AugmentCRC(initValue))
End Function


Public Function LShift(mynum As Long, mShift As Integer) As Long
Dim myBin As String
Dim myLen As Integer
myBin = dec2bin(mynum)
myLen = Len(myBin)
myBin = myBin & String(mShift, "0")
If Len(myBin) > 14 Then myBin = Mid(myBin, 2, myLen)
LShift = bin2dec(myBin)
End Function

Public Function LShift2(mynum As Long, mShift As Integer) As Long
Dim myBin As String
Dim myLen As Long
myBin = dec2bin(mynum)
myLen = Len(myBin)
myBin = myBin & String(mShift, "0")
If Len(myBin) > 16 Then myBin = Right(myBin, 16)
LShift2 = bin2dec(myBin)
End Function

Public Function RShift3(mynum As Long, mShift As Integer) As Long
Dim tempNum As Double
tempNum = mynum / (2 ^ mShift)
tempNum = Int(tempNum)
RShift3 = tempNum
End Function

Public Function LShift3(mynum As Long, mShift As Integer) As Long
Dim tempNum As Double
tempNum = mynum * (2 ^ mShift)
tempNum = Int(tempNum)
LShift3 = tempNum
End Function

Public Function RShift(mynum As Long, mShift As Integer) As Long
Dim myBin As String
Dim myLen As Integer
myBin = dec2bin(mynum)
myLen = Len(myBin)
myBin = String(mShift, "0") & myBin
myBin = Mid(myBin, 1, myLen)
RShift = bin2dec(myBin)
End Function

Private Function AugmentCRC(myCRC As Long) As Long
Dim ctr As Integer
Dim XorFlag As Boolean
XorFlag = False
For ctr = 1 To 16
    If myCRC And 32768 Then
        XorFlag = True
    Else
        XorFlag = False
    End If
    myCRC = LShift(myCRC, 1)
    If XorFlag Then myCRC = myCRC Xor 4129
Next ctr
AugmentCRC = myCRC
End Function

Public Function showCRC(myStr As String) As String
Dim tempStr As String
tempStr = Mid(myStr, 4)
tempStr = Mid(tempStr, 1, Len(tempStr) - 5)
showCRC = "CRC16 CCITT: " & CRC16_CCITT(tempStr) & vbCrLf & _
          "CRC16 CCITT2: " & CRC16_CCITT2(tempStr) & vbCrLf & _
          myStr
End Function
