Attribute VB_Name = "mduUtf8"
Public Function UTF8Decode(ByVal code As String) As String
    If code = "" Then
        UTF8Decode = ""
        Exit Function
    End If
    
    Dim tmp As String
    Dim decodeStr As String
    Dim codelen As Long
    Dim result As String
    Dim leftStr As String
     
    leftStr = Left(code, 1)
     
    While (code <> "")
        codelen = Len(code)
        leftStr = Left(code, 1)
        If leftStr = "%" Then
                If (Mid(code, 2, 1) = "C" Or Mid(code, 2, 1) = "B") Then
                    decodeStr = Replace(Mid(code, 1, 6), "%", "")
                    tmp = c10ton(Val("&H" & Hex(Val("&H" & decodeStr) And &H1F3F)))
                    tmp = String(16 - Len(tmp), "0") & tmp
                    UTF8Decode = UTF8Decode & UTF8Decode & ChrW(Val("&H" & c2to16(Mid(tmp, 3, 4)) & c2to16(Mid(tmp, 7, 2) & Mid(tmp, 11, 2)) & Right(decodeStr, 1)))
                    code = Right(code, codelen - 6)
                ElseIf (Mid(code, 2, 1) = "E") Then
                    decodeStr = Replace(Mid(code, 1, 9), "%", "")
                    tmp = c10ton((Val("&H" & Mid(Hex(Val("&H" & decodeStr) And &HF3F3F), 2, 3))))
                    tmp = String(10 - Len(tmp), "0") & tmp
                    UTF8Decode = UTF8Decode & ChrW(Val("&H" & (Mid(decodeStr, 2, 1) & c2to16(Mid(tmp, 1, 4)) & c2to16(Mid(tmp, 5, 2) & Right(tmp, 2)) & Right(decodeStr, 1))))
                    code = Right(code, codelen - 9)
                End If
        Else
            UTF8Decode = UTF8Decode & leftStr
            code = Right(code, codelen - 1)
        End If
    Wend
End Function
Public Function c2to16(ByVal x As String) As String
   Dim i As Long
   i = 1
   For i = 1 To Len(x) Step 4
      c2to16 = c2to16 & Hex(c2to10(Mid(x, i, 4)))
   Next
End Function
 
'二进制代码转换为十进制代码
Public Function c2to10(ByVal x As String) As String
   c2to10 = 0
   If x = "0" Then Exit Function
   Dim i As Long
   i = 0
   For i = 0 To Len(x) - 1
      If Mid(x, Len(x) - i, 1) = "1" Then c2to10 = c2to10 + 2 ^ (i)
   Next
End Function
 
'10进制转n进制(默认2)
Public Function c10ton(ByVal x As Integer, Optional ByVal n As Integer = 2) As String
    Dim i As Integer
    i = x \ n
    If i > 0 Then
        If x Mod n > 10 Then
            c10ton = c10ton(i, n) + Chr(x Mod n + 55)
        Else
            c10ton = c10ton(i, n) + CStr(x Mod n)
        End If
    Else
        If x > 10 Then
            c10ton = Chr(x + 55)
        Else
            c10ton = CStr(x)
        End If
    End If
End Function
