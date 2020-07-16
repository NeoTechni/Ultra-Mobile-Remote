Attribute VB_Name = "exeInfo"
Public Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Public Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal Path As String, ByVal cbBytes As Long) As Long
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
    
Public Type FILEPROPERTIE
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OrigionalFileName As String
    ProductName As String
    ProductVersion As String
End Type

Public Function ParseEnvirons(ByVal Path As String) As String
    Dim tempstr As String, temp As Long, Start As Long, Finish As Long, Env As String
    Start = InStr(Path, "%")
    Do While Start > 0
        Finish = InStr(Start + 1, Path, "%")
        If Finish <= Start Then
            Start = 0
        Else
            tempstr = Mid(Path, Start, Finish - Start + 1)
            Env = Environ(Mid(tempstr, 2, Len(tempstr) - 2))
            If Len(Env) = 0 Then
                Start = Finish
            Else
                Path = Replace(Path, tempstr, Env)
                Start = Start + Len(Env) + 1
            End If
        End If
    Loop
    ParseEnvirons = Path
End Function
Public Function EXEname(ByVal Path As String) As String
    Dim temp As FILEPROPERTIE, temp2 As Long, tempstr As String
    For temp2 = 1 To 9
        Path = Replace(Path, "%" & temp2, Empty)
    Next
    Path = Replace(Path, """", Empty)
    If InStr(Path, "/") Then Path = Left(Path, InStr(Path, "/") - 1)
    Path = Trim(Path)
    Path = ParseEnvirons(Path)
    
    exeFileInfo Path, temp
    
    If Len(temp.ProductName) = 0 Then
        tempstr = Right(Path, Len(Path) - InStrRev(Path, "\"))
        tempstr = Replace(tempstr, ".exe", Empty, , , vbTextCompare)
        EXEname = tempstr
    Else
        If StrComp(temp.ProductName, "Microsoft® Windows® Operating System", vbTextCompare) = 0 Then
            EXEname = temp.FileDescription
        Else
            EXEname = temp.ProductName
        End If
    End If
End Function

Public Function exeFileInfo(ByVal PathWithFilename As String, filedata As FILEPROPERTIE)
 ' return file-properties of given file  (EXE , DLL , OCX)
Static BACKUP As FILEPROPERTIE   ' backup info for next call without filename
If Len(PathWithFilename) = 0 Then
    filedata = BACKUP
    Exit Function
End If

Dim lngBufferlen As Long
Dim lngDummy As Long
Dim lngRc As Long
Dim lngVerPointer As Long
Dim lngHexNumber As Long
Dim bytBuffer() As Byte
Dim bytBuff(255) As Byte
Dim strBuffer As String
Dim strLangCharset As String
Dim strVersionInfo(7) As String
Dim strTemp As String
Dim intTemp As Integer
       
' size
lngBufferlen = GetFileVersionInfoSize(PathWithFilename, lngDummy)
If lngBufferlen > 0 Then
   ReDim bytBuffer(lngBufferlen)
   lngRc = GetFileVersionInfo(PathWithFilename, 0&, lngBufferlen, bytBuffer(0))
   If lngRc <> 0 Then
      lngRc = VerQueryValue(bytBuffer(0), "\VarFileInfo\Translation", _
               lngVerPointer, lngBufferlen)
      If lngRc <> 0 Then
         'lngVerPointer is a pointer to four 4 bytes of Hex number,
         'first two bytes are language id, and last two bytes are code
         'page. However, strLangCharset needs a  string of
         '4 hex digits, the first two characters correspond to the
         'language id and last two the last two character correspond
         'to the code page id.
         MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
         lngHexNumber = bytBuff(2) + bytBuff(3) * &H100 + _
                bytBuff(0) * &H10000 + bytBuff(1) * &H1000000
         strLangCharset = Hex(lngHexNumber)
         'now we change the order of the language id and code page
         'and convert it into a string representation.
         'For example, it may look like 040904E4
         'Or to pull it all apart:
         '04------        = SUBLANG_ENGLISH_USA
         '--09----        = LANG_ENGLISH
         ' ----04E4 = 1252 = Codepage for Windows:Multilingual
         Do While Len(strLangCharset) < 8
             strLangCharset = "0" & strLangCharset
         Loop
         ' assign propertienames
         strVersionInfo(0) = "CompanyName"
         strVersionInfo(1) = "FileDescription"
         strVersionInfo(2) = "FileVersion"
         strVersionInfo(3) = "InternalName"
         strVersionInfo(4) = "LegalCopyright"
         strVersionInfo(5) = "OriginalFileName"
         strVersionInfo(6) = "ProductName"
         strVersionInfo(7) = "ProductVersion"
         ' loop and get fileproperties
         For intTemp = 0 To 7
            strBuffer = String$(255, 0)
            strTemp = "\StringFileInfo\" & strLangCharset _
               & "\" & strVersionInfo(intTemp)
            lngRc = VerQueryValue(bytBuffer(0), strTemp, _
                  lngVerPointer, lngBufferlen)
            If lngRc <> 0 Then
               ' get and format data
               lstrcpy strBuffer, lngVerPointer
               strBuffer = Mid$(strBuffer, 1, InStr(strBuffer, Chr(0)) - 1)
               strVersionInfo(intTemp) = strBuffer
             Else
               ' property not found
               strVersionInfo(intTemp) = "?"
            End If
         Next intTemp
      End If
   End If
End If
' assign array to user-defined-type
With filedata
.CompanyName = strVersionInfo(0)
.FileDescription = strVersionInfo(1)
.FileVersion = strVersionInfo(2)
.InternalName = strVersionInfo(3)
.LegalCopyright = strVersionInfo(4)
.OrigionalFileName = strVersionInfo(5)
.ProductName = strVersionInfo(6)
.ProductVersion = strVersionInfo(7)
End With
BACKUP = filedata
End Function



