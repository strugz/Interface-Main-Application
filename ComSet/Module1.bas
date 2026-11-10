Attribute VB_Name = "Module1"
Public Function getSTKsData(mess As String, searchTag As String, DefaultValue As String, Optional colCount As Byte = 1) As String
Dim tempStr As String
Dim tempColCount As Integer
Dim StartCount As Long
Dim foundData As Boolean
Dim dataEnd As Boolean
foundData = False
dataEnd = False
tempStr = ""
tempColCount = 0
If IsNumeric(colCount) = False Then colCount = 1
StartCount = InStr(1, mess, searchTag)
If StartCount > 0 Then
    StartCount = StartCount + Len(searchTag)
    Do While tempColCount <> colCount And StartCount <= Len(mess)
        Do While (foundData = False Or dataEnd = False)
            If dataEnd = False Then tempStr = tempStr & Mid(mess, StartCount, 1)
            If foundData = True Then
                If Mid(mess, StartCount, 1) = " " Or Mid(mess, StartCount, 1) = vbCrLf Or Mid(mess, StartCount, 1) = vbCr Or StartCount >= Len(mess) Then dataEnd = True
            Else
                If Len(Trim(Mid(mess, StartCount, 1))) > 0 Then
                    foundData = True
                Else
                    If StartCount = Len(mess) Then
                        foundData = True
                        dataEnd = True
                    End If
                End If
            End If
            StartCount = StartCount + 1
        Loop
        tempColCount = tempColCount + 1
        If tempColCount <> colCount Then
            tempStr = ""
            foundData = False
            dataEnd = False
        Else
            tempStr = Trim(tempStr)
        End If
    Loop
Else
    tempStr = DefaultValue
End If
If Len(Trim(tempStr)) = 0 Then tempStr = DefaultValue
getSTKsData = tempStr
End Function
