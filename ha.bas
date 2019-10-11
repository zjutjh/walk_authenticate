Attribute VB_Name = "ModuleMain"
Public Declare Function CVR_InitComm Lib "termb.dll" (ByVal Port As Long) As Integer
Public Declare Function CVR_CloseComm Lib "termb.dll" () As Integer
Public Declare Function CVR_Authenticate Lib "termb.dll" () As Integer
Public Declare Function CVR_Read_Content Lib "termb.dll" (ByVal Active As Long) As Integer
Public Declare Function CVR_Ant Lib "termb.dll" (ByVal mode As Long) As Integer

Public Declare Function GetPeopleName Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer
Public Declare Function GetPeopleAddress Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer
Public Declare Function GetPeopleIDCode Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Public Declare Function SendMessage Lib "user32" _
            Alias "SendMessageA" (ByVal hwnd As Long, _
            ByVal wMsg As Long, ByVal wParam As Long, _
            lParam As Any) As Long
Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Public Const CP_UTF8 = 65001
Public Function DecodeUTF8(ByVal sUtf8 As String) As String
On Error GoTo hError
   Dim lngUtf8Size      As Long
   Dim strBuffer        As String
   Dim lngBufferSize    As Long
   Dim lngResult        As Long
   Dim bytUtf8()        As Byte
   Dim n                As Long

   If LenB(sUtf8) = 0 Then Exit Function
      Debug.Print LenB(sUtf8)
      bytUtf8 = StrConv(sUtf8, vbFromUnicode)
      lngUtf8Size = UBound(bytUtf8) + 1
      On Error GoTo 0
      'Set buffer for longest possible string i.e. each byte is
      'ANSI, thus 1 unicode(2 bytes)for every utf-8 character.
      lngBufferSize = lngUtf8Size * 2
      strBuffer = String$(lngBufferSize, vbNullChar)
      'Translate using code page 65001(UTF-8)
      lngResult = MultiByteToWideChar(CP_UTF8, 0, bytUtf8(0), lngUtf8Size, _
                    StrPtr(strBuffer), lngBufferSize)
      'Trim result to actual length
      If lngResult Then
         DecodeUTF8 = Left$(strBuffer, lngResult)
      End If
hFunEnd:
    Exit Function
hError:

End Function


