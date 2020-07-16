Attribute VB_Name = "Registryhandling"
Option Explicit

Private Const ERROR_SUCCESS = 0&
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Const KEY_ALL_ACCESS = &HF003F
    Const HKEY_DYN_DATA = &H80000006
    Const REG_BINARY = 3
    Const REG_DWORD = 4
    Const REG_DWORD_BIG_ENDIAN = 5
    Const REG_DWORD_LITTLE_ENDIAN = 4
    Const REG_EXPAND_SZ = 2
    Const REG_LINK = 6
    Const REG_MULTI_SZ = 7
    Const REG_NONE = 0
    Const REG_RESOURCE_LIST = 8
    Const REG_SZ = 1
    Const REG_FULL_RESOURCE_DESCRIPTOR = 9
    Const REG_RESOURCE_REQUIREMENTS_LIST = 10



Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
    Const Spi_seticons As Integer = 88
Dim R As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long

Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String)
On Error Resume Next
    Dim lResult As Long
    Dim lValueType As Long
    Dim strBuf As String
    Dim lDataBufSize As Long, ERROR_SUCCESS As Long
    On Error GoTo 0
    lResult = RegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Or lValueType = REG_EXPAND_SZ Then
            strBuf = String(lDataBufSize, " ")
            lResult = RegQueryValueEx(hKey, strValueName, 0&, 0&, ByVal strBuf, lDataBufSize)
            If lResult = ERROR_SUCCESS Then RegQueryStringValue = StripTerminator(strBuf)
        End If
    End If
End Function

Public Function GetString(hKey As Long, strpath As String, Optional strvalue As String, Optional Default As String = Empty)
On Error Resume Next
    Dim keyhand&, temp As String
    Dim datatype&
    R = RegOpenKey(hKey, strpath, keyhand&)
    temp = RegQueryStringValue(keyhand&, strvalue)
    If temp = Empty Then GetString = Default Else GetString = temp
    R = RegCloseKey(keyhand&)
End Function

Function StripTerminator(ByVal strString As String) As String
On Error Resume Next
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Public Sub SaveString(hKey As Long, strpath As String, strvalue As String, strdata As String)
On Error Resume Next
    Dim keyhand&
    R = RegCreateKey(hKey, strpath, keyhand&)
    R = RegSetValueEx(keyhand&, strvalue, 0, REG_SZ, ByVal strdata, Len(strdata))
    R = RegCloseKey(keyhand&)
End Sub

Public Sub Delstring(hKey As Long, strpath As String, sKey As String)
On Error Resume Next
    Dim keyhand&
    R = RegOpenKey(hKey, strpath, keyhand&)
    R = RegDeleteValue(keyhand&, sKey)
    R = RegCloseKey(keyhand&)
End Sub

Public Function ShellFolder(Optional Foldername As String = "Personal") As String
    ShellFolder = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", Foldername)
End Function


Public Function FileExtention(ByVal Filename As String) As String
    If InStr(Filename, "\") > 0 Then Filename = Right(Filename, Len(Filename) - InStrRev(Filename, "\"))
    If InStr(Filename, ".") > 0 Then Filename = Right(Filename, Len(Filename) - InStrRev(Filename, "."))
    FileExtention = Filename
End Function
Public Function FileClassName(Extention As String, Optional Default As String) As String
    FileClassName = GetString(HKEY_CLASSES_ROOT, "." & FileExtention(Extention), , Default)
End Function
Public Function FileTypeName(Extention As String, Optional Default As String, Optional Suffix As String = " file") As String
    Dim temp As String
    temp = FileClassName(Extention)
    If Len(Default) = 0 Then
        If InStr(Suffix, "*") > 0 Then
            Default = Replace(Suffix, "*", Extention)
        Else
            Default = Extention & Suffix
        End If
    End If
    If Len(temp) > 0 Then temp = GetString(HKEY_CLASSES_ROOT, temp, , Default) Else temp = Default
    FileTypeName = temp
End Function


Public Function EXEPATH()
    EXEPATH = Replace(App.Path & "\" & App.EXEName & ".exe", "\\", "\")
End Function





Public Function EnumRegKeys(ByVal Section As Long, ByVal key_name As String, List, Count As Long, Optional MarkSubkeys As Boolean) As Long
    'Dim subkeys As Collection, subkey_values As Collection
    Dim subkey_num As Integer, i As Integer, value_data(1 To 1024) As Byte, txt As String, value_data_len As Long
    Dim subkey_name As String, subkey_value As String, subkey_txt As String, value_name As String, value_string As String
    Dim Length As Long, hKey As Long, value_num As Long, value_name_len As Long, reserved As Long, value_type As Long

    'Set subkeys = New Collection
    'Set subkey_values = New Collection
    
    ' Open the key.
    If RegOpenKeyEx(Section, key_name, 0&, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then Exit Function

    ' Enumerate the key's values.
    value_num = 0
    Do
        value_name_len = 1024
        value_name = SPACE$(value_name_len)
        value_data_len = 1024

        If RegEnumValue(hKey, value_num, value_name, value_name_len, 0, value_type, value_data(1), value_data_len) <> ERROR_SUCCESS Then Exit Do

        value_name = Left$(value_name, value_name_len)

        Count = Count + 1
        ReDim Preserve List(0 To 1, Count)
        List(0, Count - 1) = value_name
        value_string = Empty
        
        Select Case value_type
            Case REG_DWORD
                value_string = "&H" & Format$(Hex$(value_data(4)), "00") & Format$(Hex$(value_data(3)), "00") & Format$(Hex$(value_data(2)), "00") & Format$(Hex$(value_data(1)), "00")
            Case REG_BINARY:                        value_string = "[binary]"
            Case REG_DWORD_BIG_ENDIAN:              value_string = "[dword big endian]"
            Case REG_DWORD_LITTLE_ENDIAN:           value_string = "[dword little endian]"
            Case REG_EXPAND_SZ:                     value_string = "[expand sz]"
            Case REG_FULL_RESOURCE_DESCRIPTOR:      value_string = "[full resource descriptor]"
            Case REG_LINK:                          value_string = "[link]"
            Case REG_MULTI_SZ:                      value_string = "[multi sz]"
            Case REG_NONE:                          value_string = "[none]"
            Case REG_RESOURCE_LIST:                 value_string = "[resource list]"
            Case REG_RESOURCE_REQUIREMENTS_LIST:    value_string = "[resource requirements list]"
            Case REG_SZ
                For i = 1 To value_data_len - 1
                    value_string = value_string & Chr$(value_data(i))
                Next i
        End Select
        List(1, Count - 1) = value_string
        value_num = value_num + 1
    Loop


    ' Enumerate the subkeys.
    subkey_num = 0
    Do
        ' Enumerate subkeys until we get an error.
        Length = 256
        subkey_name = SPACE$(Length)
        If RegEnumKey(hKey, subkey_num, subkey_name, Length) <> ERROR_SUCCESS Then Exit Do
        subkey_num = subkey_num + 1
        
        subkey_name = Left$(subkey_name, InStr(subkey_name, Chr$(0)) - 1)
        'subkeys.Add subkey_name
    
        Count = Count + 1
        ReDim Preserve List(0 To 1, Count)
        List(0, Count - 1) = subkey_name
    
        ' Get the subkey's value.
        Length = 256
        subkey_value = SPACE$(Length)
        If RegQueryValue(hKey, subkey_name, subkey_value, Length) = ERROR_SUCCESS Then
            ' Remove the trailing null character.
            subkey_value = Left$(subkey_value, Length - 1)
            
            If MarkSubkeys Then subkey_value = "[subkey]"
            'subkey_values.Add subkey_value
            List(1, Count - 1) = subkey_value
        End If
        
        EnumRegKeys Section, key_name & "\" & subkey_name, List, Count, MarkSubkeys
    Loop
    
    ' Close the key.
    RegCloseKey hKey

    EnumRegKeys = Count
End Function

Public Function EnumVerbs(Extention As String, List, Optional IgnoreOpen As Boolean = True) As Long
    Dim KeyList() As String, KeyCount As Long, temp As Long, Count As Long, doit As Boolean, Check As Boolean
    If InStr(Extention, ".") = 0 Then
        EnumRegKeys HKEY_CLASSES_ROOT, FileClassName(Extention) & "\shell", KeyList, KeyCount, True
    Else
        EnumRegKeys HKEY_CLASSES_ROOT, "Applications\" & Extention & "\shell", KeyList, KeyCount, True
        Check = True
    End If
    
    For temp = 0 To KeyCount - 1
        If KeyList(1, temp) = "[subkey]" Then
            doit = True
            
            Select Case LCase(KeyList(0, temp))
                Case "command", "defaulticon", "supportedtypes", "droptarget", "taskbarexception": doit = False
                Case "open": If IgnoreOpen Then doit = False
            End Select
            
            If Check And doit Then
                doit = Len(GetString(HKEY_CLASSES_ROOT, "Applications\" & Extention & "\shell\" & KeyList(0, temp) & "\command")) > 0
            End If
            
            If doit Then
                Count = Count + 1
                ReDim Preserve List(Count)
                List(Count - 1) = KeyList(0, temp)
            End If
        End If
    Next
    EnumVerbs = Count
End Function

Public Function EnumOpenWith(Extention As String, List) As Long
    'HKEY_CURRENT_USER Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.xyz\OpenWithProgIDs
    'HKEY_CLASSES_ROOT\.xyz\OpenWithList
    'HKEY_CLASSES_ROOT\.xyz\OpenWithProgIDs
    'HKEY_CLASSES_ROOT\SystemFileAssociations\<Perceived Type>\OpenWithList
    
    Dim Count As Long, temp As Long
    Count = EnumMRUlist(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\." & Extention & "\OpenWithList", List)
        
    EnumOpenWith = Count
End Function

Public Function FindProgram(Name As String, Optional ByVal CMD As String) As String
    Dim tempstr As String
    tempstr = GetString(HKEY_CLASSES_ROOT, "Applications\" & Name & "\shell\open\command")
    If Len(CMD) > 0 Then
        If Left(CMD, 1) <> """" Then CMD = """" & CMD
        If Right(CMD, 1) <> """" Then CMD = CMD & """"
        tempstr = Replace(tempstr, "%1", CMD)
    End If
    FindProgram = tempstr
End Function

Public Function EnumPrograms(List) As Long
    Dim temp As Long, templist() As String, Count As Long, count2 As Long
    EnumRegKeys HKEY_CLASSES_ROOT, "Applications", templist, Count, True
    For temp = 0 To Count - 1
        If templist(1, temp) = "[subkey]" Then
            If StrComp(GetExtention(templist(0, temp)), "exe", vbTextCompare) = 0 Then
                count2 = count2 + 1
                ReDim Preserve List(count2)
                List(count2 - 1) = templist(0, temp)
            End If
        End If
    Next
    EnumPrograms = count2
End Function

Public Function GetVerbPath(Application As String, Optional Verb As String = "Open") As String
    GetVerbPath = GetString(HKEY_CLASSES_ROOT, "Applications\" & Application & "\shell\" & Verb & "\command", Empty)
End Function

Public Function EnumMRUlist(hKey As Long, strpath As String, List, Optional sKey As String = "MRUList") As Long
    Dim temp As Long, Count As Long, MRUlist As String
    MRUlist = GetString(hKey, strpath, sKey)
    Count = Len(MRUlist)
    If Count > 0 Then
        ReDim List(Count)
        For temp = 1 To Count
            List(temp - 1) = GetString(hKey, strpath, Mid(MRUlist, temp, 1))
        Next
    End If
    EnumMRUlist = Count
End Function

'Public Sub Test()
'    Dim List() As String, Count As Long, temp As Long
'    Count = EnumVerbs("alcohol.exe", List, False) ' EnumOpenWith("torrent", List)
'    For temp = 0 To Count - 1
'        Debug.Print List(temp)
'    Next
'End Sub
